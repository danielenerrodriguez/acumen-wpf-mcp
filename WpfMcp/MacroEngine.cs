using System.IO;
using System.Text;
using System.Text.Json;
using System.Text.RegularExpressions;

namespace WpfMcp;

/// <summary>
/// Loads, validates, and executes macro YAML files.
/// Macros are discovered from the macros/ folder next to the exe,
/// or from the path specified by WPFMCP_MACROS_PATH env var.
/// Watches the macros directory for changes and auto-reloads.
/// Also loads _knowledge.yaml files as knowledge bases for AI agents.
/// </summary>
public class MacroEngine : IDisposable
{
    private readonly Dictionary<string, MacroDefinition> _macros = new(StringComparer.OrdinalIgnoreCase);
    private readonly List<MacroLoadError> _loadErrors = new();
    private readonly Dictionary<string, KnowledgeBase> _knowledgeBases = new(StringComparer.OrdinalIgnoreCase);
    /// <summary>Cached parsed dictionaries for knowledge bases, populated during Reload().</summary>
    private readonly Dictionary<string, Dictionary<string, object>> _parsedKnowledgeBases = new(StringComparer.OrdinalIgnoreCase);
    private readonly string _macrosPath;
    private FileSystemWatcher? _watcher;
    private Timer? _debounceTimer;
    private readonly object _reloadLock = new();
    private bool _disposed;

    /// <summary>Fires after macros/knowledge bases are reloaded (from FileSystemWatcher or manual Reload call).</summary>
    public event Action? OnReloaded;

    // Use shared deserializers from YamlHelpers to avoid duplicate configuration

    private static readonly Regex s_paramPattern = new(@"\{\{(\w+)\}\}", RegexOptions.Compiled);

    /// <summary>Root path of the macros directory.</summary>
    public string MacrosPath => _macrosPath;

    /// <summary>YAML files that failed to parse on the last reload.</summary>
    public IReadOnlyList<MacroLoadError> LoadErrors
    {
        get { lock (_reloadLock) return _loadErrors.ToList(); }
    }

    /// <summary>Knowledge base files loaded from the macros directory.</summary>
    public IReadOnlyList<KnowledgeBase> KnowledgeBases
    {
        get { lock (_reloadLock) return _knowledgeBases.Values.ToList(); }
    }

    public MacroEngine(string? macrosPath = null, bool enableWatcher = true)
    {
        _macrosPath = Constants.ResolveMacrosPath(macrosPath);
        Reload();

        if (enableWatcher && Directory.Exists(_macrosPath))
            StartWatcher();
    }

    /// <summary>Scan the macros directory and load all .yaml files.</summary>
    public void Reload()
    {
        lock (_reloadLock)
        {
            _macros.Clear();
            _loadErrors.Clear();
            _knowledgeBases.Clear();
            _parsedKnowledgeBases.Clear();
            if (!Directory.Exists(_macrosPath)) return;

            foreach (var file in Directory.GetFiles(_macrosPath, "*.yaml", SearchOption.AllDirectories))
            {
                var fileName = Path.GetFileName(file);
                var relativePath = Path.GetRelativePath(_macrosPath, file);

                // Knowledge base files use underscore prefix convention (_knowledge.yaml)
                if (fileName.StartsWith('_'))
                {
                    LoadKnowledgeBase(file, relativePath);
                    continue;
                }

                var macroName = relativePath
                    .Replace(Path.DirectorySeparatorChar, '/')
                    .Replace(".yaml", "", StringComparison.OrdinalIgnoreCase);

                try
                {
                    var yaml = File.ReadAllText(file);
                    var macro = YamlHelpers.Deserializer.Deserialize<MacroDefinition>(yaml);
                    if (macro == null)
                    {
                        _loadErrors.Add(new MacroLoadError(relativePath, macroName,
                            "YAML deserialized to null (empty or invalid file)"));
                        continue;
                    }

                    if (macro.Steps.Count == 0)
                    {
                        _loadErrors.Add(new MacroLoadError(relativePath, macroName,
                            "Macro has no steps defined"));
                        continue;
                    }

                    _macros[macroName] = macro;
                }
                catch (Exception ex)
                {
                    _loadErrors.Add(new MacroLoadError(relativePath, macroName, ex.Message));
                    Console.Error.WriteLine($"Warning: Failed to load macro '{file}': {ex.Message}");
                }
            }

            // Expand include steps — must happen after all macros are loaded
            ExpandIncludes();

            var parts = new List<string>();
            if (_macros.Count > 0) parts.Add($"{_macros.Count} macros");
            if (_knowledgeBases.Count > 0) parts.Add($"{_knowledgeBases.Count} knowledge base(s)");
            if (_loadErrors.Count > 0) parts.Add($"{_loadErrors.Count} errors");
            if (parts.Count > 0)
                Console.Error.WriteLine($"[{DateTime.Now:HH:mm:ss}] Reloaded: {string.Join(", ", parts)}");
        }

        // Fire outside lock to prevent deadlocks with subscribers
        OnReloaded?.Invoke();
    }

    /// <summary>
    /// Expand all 'action: include' steps in loaded macros by inlining the referenced
    /// macro's steps at load time. This produces flat step lists with zero runtime overhead.
    /// Detects circular includes and reports them as load errors.
    /// </summary>
    private void ExpandIncludes()
    {
        // Process each macro that contains include steps
        foreach (var (macroName, macro) in _macros.ToList())
        {
            if (!macro.Steps.Any(s => string.Equals(s.Action, "include", StringComparison.OrdinalIgnoreCase)))
                continue;

            var expanded = ExpandMacroSteps(macroName, macro, new HashSet<string>(StringComparer.OrdinalIgnoreCase) { macroName });
            if (expanded != null)
                macro.Steps = expanded;
        }
    }

    /// <summary>
    /// Recursively expand include steps for a macro. Returns the expanded step list,
    /// or null if an error was recorded (circular ref, missing macro).
    /// </summary>
    private List<MacroStep>? ExpandMacroSteps(string macroName, MacroDefinition macro, HashSet<string> visited)
    {
        var result = new List<MacroStep>();

        foreach (var step in macro.Steps)
        {
            if (!string.Equals(step.Action, "include", StringComparison.OrdinalIgnoreCase))
            {
                result.Add(step);
                continue;
            }

            var includedName = step.MacroName;
            if (string.IsNullOrEmpty(includedName))
            {
                _loadErrors.Add(new MacroLoadError("", macroName,
                    "include step is missing 'macro_name' field"));
                return null;
            }

            if (visited.Contains(includedName))
            {
                _loadErrors.Add(new MacroLoadError("", macroName,
                    $"Circular include detected: {macroName} -> {includedName}"));
                return null;
            }

            if (!_macros.TryGetValue(includedName, out var includedMacro))
            {
                _loadErrors.Add(new MacroLoadError("", macroName,
                    $"include references unknown macro '{includedName}'"));
                return null;
            }

            // Recursively expand the included macro's steps first
            visited.Add(includedName);
            var childSteps = ExpandMacroSteps(includedName, includedMacro, visited);
            visited.Remove(includedName);

            if (childSteps == null) return null; // Error already recorded

            // Clone steps and apply parameter remapping if params are specified
            var paramMap = step.Params; // e.g., { "filePath": "{{inputFile}}" }
            foreach (var childStep in childSteps)
            {
                var cloned = CloneStepWithParamRemap(childStep, paramMap);
                result.Add(cloned);
            }
        }

        return result;
    }

    /// <summary>
    /// Clone a MacroStep, remapping parameter references in all string fields.
    /// If paramMap is { "childParam": "{{parentParam}}" }, then all occurrences of
    /// {{childParam}} in the cloned step's string properties are replaced with {{parentParam}}.
    /// This enables parent macros to pass their parameters through to included macros.
    /// </summary>
    private static MacroStep CloneStepWithParamRemap(MacroStep source, Dictionary<string, string>? paramMap)
    {
        var clone = new MacroStep
        {
            Action = source.Action,
            AutomationId = source.AutomationId,
            Name = source.Name,
            ClassName = source.ClassName,
            ControlType = source.ControlType,
            Path = source.Path != null ? new List<string>(source.Path) : null,
            SaveAs = source.SaveAs,
            Ref = source.Ref,
            Text = source.Text,
            Value = source.Value,
            Keys = source.Keys,
            MaxDepth = source.MaxDepth,
            ProcessName = source.ProcessName,
            Pid = source.Pid,
            Seconds = source.Seconds,
            MacroName = source.MacroName,
            Params = source.Params != null ? new Dictionary<string, string>(source.Params) : null,
            ExePath = source.ExePath,
            Arguments = source.Arguments,
            WorkingDirectory = source.WorkingDirectory,
            IfNotRunning = source.IfNotRunning,
            TitleContains = source.TitleContains,
            Enabled = source.Enabled,
            Property = source.Property,
            Expected = source.Expected,
            Message = source.Message,
            MatchMode = source.MatchMode,
            StepTimeout = source.StepTimeout,
            RetryInterval = source.RetryInterval,
        };

        if (paramMap == null || paramMap.Count == 0)
            return clone;

        // Build a reverse map: {{childParam}} -> replacement value from parent
        // e.g., paramMap = { "filePath": "{{inputFile}}" }
        // means replace {{filePath}} with {{inputFile}} in all string fields
        clone.AutomationId = RemapParam(clone.AutomationId, paramMap);
        clone.Name = RemapParam(clone.Name, paramMap);
        clone.ClassName = RemapParam(clone.ClassName, paramMap);
        clone.ControlType = RemapParam(clone.ControlType, paramMap);
        clone.Text = RemapParam(clone.Text, paramMap);
        clone.Value = RemapParam(clone.Value, paramMap);
        clone.Keys = RemapParam(clone.Keys, paramMap);
        clone.ExePath = RemapParam(clone.ExePath, paramMap);
        clone.Arguments = RemapParam(clone.Arguments, paramMap);
        clone.WorkingDirectory = RemapParam(clone.WorkingDirectory, paramMap);
        clone.TitleContains = RemapParam(clone.TitleContains, paramMap);
        clone.Expected = RemapParam(clone.Expected, paramMap);
        clone.Message = RemapParam(clone.Message, paramMap);
        clone.MacroName = RemapParam(clone.MacroName, paramMap);
        clone.ProcessName = RemapParam(clone.ProcessName, paramMap);
        clone.Ref = RemapParam(clone.Ref, paramMap);
        clone.SaveAs = RemapParam(clone.SaveAs, paramMap);
        if (clone.Path != null)
        {
            for (int i = 0; i < clone.Path.Count; i++)
                clone.Path[i] = RemapParam(clone.Path[i], paramMap) ?? clone.Path[i];
        }
        if (clone.Params != null)
        {
            var remapped = new Dictionary<string, string>();
            foreach (var (k, v) in clone.Params)
                remapped[k] = RemapParam(v, paramMap) ?? v;
            clone.Params = remapped;
        }

        return clone;
    }

    /// <summary>
    /// Replace {{childParam}} references with the mapped value from the include's params.
    /// e.g., if paramMap = { "filePath": "{{inputFile}}" }, then "{{filePath}}" becomes "{{inputFile}}".
    /// Literal values work too: if paramMap = { "filePath": "C:\\data\\test.xer" }, then
    /// "{{filePath}}" becomes "C:\\data\\test.xer".
    /// </summary>
    private static string? RemapParam(string? value, Dictionary<string, string> paramMap)
    {
        if (value == null) return null;
        foreach (var (childParam, replacement) in paramMap)
            value = value.Replace($"{{{{{childParam}}}}}", replacement, StringComparison.OrdinalIgnoreCase);
        return value;
    }

    /// <summary>Load a _knowledge.yaml file, parse it as a dictionary, and build a summary.</summary>
    private void LoadKnowledgeBase(string filePath, string relativePath)
    {
        try
        {
            var yaml = File.ReadAllText(filePath);
            var dict = YamlHelpers.DictDeserializer.Deserialize<Dictionary<string, object>>(yaml);
            if (dict == null) return;

            // Verify it has kind: knowledge-base
            if (!dict.TryGetValue("kind", out var kindObj) ||
                kindObj?.ToString() != "knowledge-base")
                return;

            // Derive product name from folder path (e.g., "acumen-fuse/_knowledge.yaml" -> "acumen-fuse")
            var dir = Path.GetDirectoryName(relativePath);
            var productName = string.IsNullOrEmpty(dir) ? "default" : dir.Replace(Path.DirectorySeparatorChar, '/');

            var summary = BuildKnowledgeSummary(dict, productName);
            _knowledgeBases[productName] = new KnowledgeBase(productName, relativePath, summary, yaml);
            _parsedKnowledgeBases[productName] = dict;
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"Warning: Failed to load knowledge base '{filePath}': {ex.Message}");
        }
    }

    /// <summary>Build a condensed text summary from the knowledge base dictionary.</summary>
    private static string BuildKnowledgeSummary(Dictionary<string, object> dict, string productName)
    {
        var sb = new StringBuilder();
        sb.AppendLine($"  {productName}:");

        // Application info
        if (dict.TryGetValue("application", out var appObj) && appObj is Dictionary<object, object> app)
        {
            var name = app.TryGetValue("name", out var n) ? n?.ToString() : "Unknown";
            var version = app.TryGetValue("version", out var v) ? v?.ToString() : "";
            var exePath = app.TryGetValue("exe_path", out var e) ? e?.ToString() : "";
            sb.AppendLine($"    Application: {name} {version}");
            if (!string.IsNullOrEmpty(exePath))
                sb.AppendLine($"    Exe Path: {exePath}");
        }

        // Installation / Samples
        if (dict.TryGetValue("installation", out var instObj) && instObj is Dictionary<object, object> inst)
        {
            var basePath = inst.TryGetValue("base_path", out var bp) ? bp?.ToString() : "";
            if (!string.IsNullOrEmpty(basePath))
                sb.AppendLine($"    Install Path: {basePath}");
            if (inst.TryGetValue("samples", out var samplesObj) && samplesObj is Dictionary<object, object> samples)
            {
                var sampleCount = 0;
                foreach (var kv in samples)
                {
                    if (kv.Value is List<object> list) sampleCount += list.Count;
                }
                if (sampleCount > 0)
                    sb.AppendLine($"    Sample Files: {sampleCount}+ files available");
            }
        }

        // Verified keytips
        if (dict.TryGetValue("keytips", out var keytipsObj) && keytipsObj is Dictionary<object, object> keytips)
        {
            if (keytips.TryGetValue("verified_sequences", out var seqObj) && seqObj is List<object> seqs)
            {
                var keytipLines = new List<string>();
                foreach (var seq in seqs)
                {
                    if (seq is Dictionary<object, object> s)
                    {
                        var keys = s.TryGetValue("keys", out var k) ? k?.ToString() : "";
                        var action = s.TryGetValue("action", out var a) ? a?.ToString() : "";
                        if (!string.IsNullOrEmpty(keys))
                            keytipLines.Add($"{keys} ({action})");
                    }
                }
                if (keytipLines.Count > 0)
                    sb.AppendLine($"    Verified Keytips: {string.Join(", ", keytipLines)}");
            }
        }

        // Keyboard shortcuts
        if (dict.TryGetValue("keyboard_shortcuts", out var shortcutsObj) && shortcutsObj is List<object> shortcuts)
        {
            var shortcutLines = new List<string>();
            foreach (var sc in shortcuts)
            {
                if (sc is Dictionary<object, object> s)
                {
                    var key = s.TryGetValue("key", out var k) ? k?.ToString() : "";
                    var action = s.TryGetValue("action", out var a) ? a?.ToString() : "";
                    if (!string.IsNullOrEmpty(key))
                        shortcutLines.Add($"{key} ({action})");
                }
            }
            if (shortcutLines.Count > 0)
                sb.AppendLine($"    Keyboard Shortcuts: {string.Join(", ", shortcutLines)}");
        }

        // Key automation IDs (just tab-level ones as highlights)
        if (dict.TryGetValue("automation_ids", out var aidsObj) && aidsObj is Dictionary<object, object> aids)
        {
            var idCount = 0;
            var tabIds = new List<string>();
            if (aids.TryGetValue("ribbon_tabs", out var tabsObj) && tabsObj is List<object> tabs)
            {
                foreach (var tab in tabs)
                {
                    if (tab is Dictionary<object, object> t && t.TryGetValue("id", out var id))
                        tabIds.Add(id.ToString()!);
                }
            }
            foreach (var section in aids.Values)
            {
                if (section is List<object> list) idCount += list.Count;
            }
            if (tabIds.Count > 0)
                sb.AppendLine($"    Key Automation IDs: {string.Join(", ", tabIds)} + {idCount} more");
        }

        // Navigation tips count
        if (dict.TryGetValue("navigation_tips", out var tipsObj) && tipsObj is List<object> tips)
            sb.AppendLine($"    Navigation Tips: {tips.Count} tips available");

        // Workflows count
        if (dict.TryGetValue("workflows", out var wfObj) && wfObj is Dictionary<object, object> workflows)
            sb.AppendLine($"    Workflows: {workflows.Count} documented workflows");

        sb.AppendLine($"    Full details: use MCP resource 'knowledge://{productName}'");

        return sb.ToString().TrimEnd();
    }

    private void StartWatcher()
    {
        _watcher = new FileSystemWatcher(_macrosPath, "*.yaml")
        {
            IncludeSubdirectories = true,
            NotifyFilter = NotifyFilters.FileName | NotifyFilters.LastWrite | NotifyFilters.CreationTime
        };

        _watcher.Created += OnFileChanged;
        _watcher.Changed += OnFileChanged;
        _watcher.Deleted += OnFileChanged;
        _watcher.Renamed += OnFileRenamed;
        _watcher.EnableRaisingEvents = true;
    }

    private void OnFileChanged(object sender, FileSystemEventArgs e) => DebouncedReload();
    private void OnFileRenamed(object sender, RenamedEventArgs e) => DebouncedReload();

    /// <summary>
    /// Debounce reload: wait 500ms of quiet before actually reloading.
    /// Prevents rapid-fire reloads when editors save multiple times.
    /// </summary>
    private void DebouncedReload()
    {
        _debounceTimer?.Dispose();
        _debounceTimer = new Timer(_ =>
        {
            try { Reload(); }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"[{DateTime.Now:HH:mm:ss}] Macro reload error: {ex.Message}");
            }
        }, null, Constants.MacroReloadDebounceMs, Timeout.Infinite);
    }

    // --- Known action types and their required fields for validation ---

    private static readonly HashSet<string> s_knownActions = new(StringComparer.OrdinalIgnoreCase)
    {
        "send_keys", "keys", "find", "find_by_path", "click", "right_click",
        "type", "set_value", "get_value", "wait", "wait_for_enabled", "macro", "include",
        "launch", "wait_for_window", "focus", "snapshot", "screenshot",
        "properties", "children", "file_dialog", "attach", "verify"
    };

    /// <summary>
    /// Validate a list of step dictionaries for known action types and required fields.
    /// Returns null if valid, or an error message string if invalid.
    /// </summary>
    public static string? ValidateSteps(List<Dictionary<string, object>> steps)
    {
        if (steps.Count == 0)
            return "Macro must have at least one step";

        for (int i = 0; i < steps.Count; i++)
        {
            var step = steps[i];
            if (!step.TryGetValue("action", out var actionObj) || actionObj is not string action || string.IsNullOrEmpty(action))
                return $"Step {i + 1}: missing 'action' field";

            var actionLower = action.ToLowerInvariant();
            if (!s_knownActions.Contains(actionLower))
                return $"Step {i + 1}: unknown action '{action}'. Valid actions: {string.Join(", ", s_knownActions.Order())}";

            var error = ValidateStepFields(i, actionLower, step);
            if (error != null) return error;
        }
        return null;
    }

    private static string? ValidateStepFields(int index, string action, Dictionary<string, object> step)
    {
        string StepPrefix() => $"Step {index + 1} ({action})";

        bool Has(string key) => step.TryGetValue(key, out var v) && v != null &&
            (v is not string s || !string.IsNullOrEmpty(s));

        switch (action)
        {
            case "send_keys" or "keys":
                if (!Has("keys")) return $"{StepPrefix()}: requires 'keys' field";
                break;
            case "find":
                if (!Has("automation_id") && !Has("name") && !Has("control_type") && !Has("class_name"))
                    return $"{StepPrefix()}: requires at least one of: automation_id, name, control_type, class_name";
                break;
            case "find_by_path":
                if (!Has("path")) return $"{StepPrefix()}: requires 'path' field";
                break;
            case "type":
                if (!Has("text")) return $"{StepPrefix()}: requires 'text' field";
                break;
            case "set_value":
                if (!Has("ref")) return $"{StepPrefix()}: requires 'ref' field";
                if (!Has("value")) return $"{StepPrefix()}: requires 'value' field";
                break;
            case "wait":
                if (!Has("seconds")) return $"{StepPrefix()}: requires 'seconds' field";
                break;
            case "macro":
            case "include":
                if (!Has("macro_name")) return $"{StepPrefix()}: requires 'macro_name' field";
                break;
            case "wait_for_window":
                if (!Has("title_contains")) return $"{StepPrefix()}: requires 'title_contains' field";
                break;
            case "wait_for_enabled":
                if (!Has("automation_id") && !Has("name") && !Has("control_type") && !Has("class_name") && !Has("ref"))
                    return $"{StepPrefix()}: requires at least one of: automation_id, name, control_type, class_name, ref";
                break;
            case "file_dialog":
                if (!Has("text")) return $"{StepPrefix()}: requires 'text' field (the file path)";
                break;
            case "attach":
                if (!Has("process_name") && !Has("pid"))
                    return $"{StepPrefix()}: requires at least one of: process_name, pid";
                break;
            case "verify":
                if (!Has("ref")) return $"{StepPrefix()}: requires 'ref' field";
                if (!Has("property")) return $"{StepPrefix()}: requires 'property' field";
                if (!Has("expected")) return $"{StepPrefix()}: requires 'expected' field";
                break;
            // click, right_click, focus, snapshot, screenshot, properties, children, launch, get_value
            // have no strictly required fields
        }
        return null;
    }

    /// <summary>
    /// Scan knowledge bases for a matching process_name field.
    /// Returns the product folder name (e.g., "acumen-fuse") or null if no match.
    /// </summary>
    public string? GetProductFolder(string processName)
    {
        lock (_reloadLock)
        {
            foreach (var (folderName, dict) in _parsedKnowledgeBases)
            {
                // Check application.process_name using the cached parsed dictionary
                if (dict.TryGetValue("application", out var appObj) && appObj is Dictionary<object, object> app)
                {
                    if (app.TryGetValue("process_name", out var pn) &&
                        string.Equals(pn?.ToString(), processName, StringComparison.OrdinalIgnoreCase))
                        return folderName;
                }
            }
        }
        return null;
    }

    /// <summary>
    /// Save a macro definition to disk as a YAML file.
    /// Steps are passed as dictionaries to produce clean YAML output (no JSON-in-YAML).
    /// Auto-derives the product folder from the attached process name via knowledge bases.
    /// </summary>
    public SaveMacroResult SaveMacro(
        string name,
        string description,
        List<Dictionary<string, object>> steps,
        List<Dictionary<string, object>>? parameters,
        int timeout,
        bool force,
        string? attachedProcessName)
    {
        // Validate steps
        var validationError = ValidateSteps(steps);
        if (validationError != null)
            return new SaveMacroResult(false, "", "", $"Validation error: {validationError}");

        // Determine product folder
        string? productFolder = null;
        if (!string.IsNullOrEmpty(attachedProcessName))
            productFolder = GetProductFolder(attachedProcessName);

        if (productFolder == null)
        {
            return new SaveMacroResult(false, "", "",
                $"Cannot determine product folder for process '{attachedProcessName}'. " +
                $"No knowledge base has a matching process_name. " +
                $"Include the product folder in the macro name (e.g., 'my-product/{name}').");
        }

        // Build the full macro name with product folder prefix (unless already included)
        var macroName = name.Contains('/')
            ? name
            : $"{productFolder}/{name}";

        // Check if macro already exists
        var relativePath = macroName.Replace('/', Path.DirectorySeparatorChar) + ".yaml";
        var fullPath = Path.Combine(_macrosPath, relativePath);

        if (File.Exists(fullPath) && !force)
        {
            return new SaveMacroResult(false, fullPath, macroName,
                $"Macro '{macroName}' already exists at {fullPath}. Use force=true to overwrite.");
        }

        // Build the YAML content as an ordered dictionary for clean output
        var yamlDoc = new Dictionary<string, object>
        {
            ["name"] = name.Contains('/') ? name.Split('/').Last() : name,
            ["description"] = description,
        };

        if (timeout > 0)
            yamlDoc["timeout"] = timeout;

        if (parameters != null && parameters.Count > 0)
            yamlDoc["parameters"] = parameters;

        yamlDoc["steps"] = steps;

        // Serialize to YAML
        var yaml = YamlHelpers.Serializer.Serialize(yamlDoc);

        // Write the file
        var dir = Path.GetDirectoryName(fullPath);
        if (!string.IsNullOrEmpty(dir) && !Directory.Exists(dir))
            Directory.CreateDirectory(dir);

        File.WriteAllText(fullPath, yaml, Encoding.UTF8);

        // FileSystemWatcher will auto-reload, but log it
        Console.Error.WriteLine($"[{DateTime.Now:HH:mm:ss}] Saved macro: {macroName} -> {fullPath}");

        return new SaveMacroResult(true, fullPath, macroName,
            $"Macro '{macroName}' saved successfully to {fullPath}");
    }

    // --- Export macros as Windows shortcuts (.lnk) ---

    /// <summary>
    /// Export a macro as a Windows shortcut (.lnk) file.
    /// The shortcut targets WpfMcp.exe with the YAML file path as argument,
    /// reusing the existing drag-and-drop execution mode.
    /// </summary>
    public ExportMacroResult ExportMacro(string macroName, string? shortcutsPath = null, bool force = false)
    {
        if (!_macros.TryGetValue(macroName, out var macro))
            return new ExportMacroResult(false, "", macroName, $"Macro '{macroName}' not found");

        var exePath = Environment.ProcessPath
            ?? Path.Combine(AppContext.BaseDirectory, "WpfMcp.exe");

        // Resolve the YAML file path (absolute)
        var yamlRelative = macroName.Replace('/', Path.DirectorySeparatorChar) + ".yaml";
        var yamlFullPath = Path.Combine(_macrosPath, yamlRelative);
        if (!File.Exists(yamlFullPath))
            return new ExportMacroResult(false, "", macroName, $"YAML file not found: {yamlFullPath}");

        // Resolve shortcuts output directory (sibling of macros folder)
        var resolvedShortcutsPath = Constants.ResolveShortcutsPath(shortcutsPath, _macrosPath);

        // Mirror the macro directory structure in shortcuts folder
        var lnkRelative = macroName.Replace('/', Path.DirectorySeparatorChar) + ".lnk";
        var lnkFullPath = Path.Combine(resolvedShortcutsPath, lnkRelative);

        if (File.Exists(lnkFullPath) && !force)
            return new ExportMacroResult(false, lnkFullPath, macroName,
                $"Shortcut already exists at {lnkFullPath}. Use force=true to overwrite.");

        try
        {
            var description = !string.IsNullOrEmpty(macro.Description)
                ? $"WPF MCP Macro: {macro.Description}"
                : $"WPF MCP Macro: {macroName}";

            // Truncate description to 260 chars (Windows .lnk limit)
            if (description.Length > 260)
                description = description[..257] + "...";

            ShortcutCreator.CreateShortcut(
                lnkPath: lnkFullPath,
                targetExe: exePath,
                arguments: $"\"{yamlFullPath}\"",
                workingDirectory: Path.GetDirectoryName(exePath) ?? "",
                description: description,
                runAsAdmin: false);

            Console.Error.WriteLine($"[{DateTime.Now:HH:mm:ss}] Exported shortcut: {macroName} -> {lnkFullPath}");
            return new ExportMacroResult(true, lnkFullPath, macroName,
                $"Shortcut exported to {lnkFullPath}");
        }
        catch (Exception ex)
        {
            return new ExportMacroResult(false, lnkFullPath, macroName,
                $"Failed to create shortcut: {ex.Message}");
        }
    }

    /// <summary>
    /// Export all loaded macros as Windows shortcuts.
    /// Returns a list of results for each macro.
    /// </summary>
    public List<ExportMacroResult> ExportAllMacros(string? shortcutsPath = null, bool force = false)
    {
        var results = new List<ExportMacroResult>();
        lock (_reloadLock)
        {
            foreach (var macroName in _macros.Keys.OrderBy(k => k))
                results.Add(ExportMacro(macroName, shortcutsPath, force));
        }
        return results;
    }

    public void Dispose()
    {
        if (_disposed) return;
        _disposed = true;
        _debounceTimer?.Dispose();
        _watcher?.Dispose();
    }

    /// <summary>List all loaded macros with their metadata.</summary>
    public List<MacroInfo> List()
    {
        var result = new List<MacroInfo>();
        foreach (var (name, macro) in _macros)
        {
            result.Add(new MacroInfo(
                name,
                macro.Name,
                macro.Description,
                macro.Parameters.Select(p => new MacroParamInfo(
                    p.Name, p.Description, p.Required, p.Default)).ToList()));
        }
        return result.OrderBy(m => m.Name).ToList();
    }

    /// <summary>Get a macro definition by name.</summary>
    public MacroDefinition? Get(string name) =>
        _macros.TryGetValue(name, out var m) ? m : null;

    /// <summary>Get the absolute file path of a macro by name, or null if not found.</summary>
    public string? GetMacroFilePath(string name)
    {
        if (!_macros.ContainsKey(name)) return null;
        var relativePath = name.Replace('/', Path.DirectorySeparatorChar) + ".yaml";
        var fullPath = Path.Combine(_macrosPath, relativePath);
        return File.Exists(fullPath) ? fullPath : null;
    }

    /// <summary>
    /// Execute a macro definition directly (for drag-drop / run-file scenarios).
    /// </summary>
    public async Task<MacroResult> ExecuteDefinitionAsync(
        MacroDefinition macro,
        string displayName,
        Dictionary<string, string>? parameters = null,
        UiaEngine? engine = null,
        ElementCache? cache = null,
        CancellationToken cancellation = default,
        Action<string>? onLog = null)
    {
        return await ExecuteInternalAsync(macro, displayName, parameters, engine, cache, cancellation, onLog);
    }

    /// <summary>
    /// Execute a macro by name with the given parameters.
    /// Uses UiaEngine directly (for elevated server / CLI mode).
    /// </summary>
    public async Task<MacroResult> ExecuteAsync(
        string macroName,
        Dictionary<string, string>? parameters = null,
        UiaEngine? engine = null,
        ElementCache? cache = null,
        CancellationToken cancellation = default,
        Action<string>? onLog = null)
    {
        if (!_macros.TryGetValue(macroName, out var macro))
            return new MacroResult(false, 0, 0, $"Macro '{macroName}' not found");

        return await ExecuteInternalAsync(macro, macroName, parameters, engine, cache, cancellation, onLog);
    }

    /// <summary>Build a human-readable summary of a macro step for logging.</summary>
    internal static string FormatStepSummary(MacroStep step, Dictionary<string, string> parameters)
    {
        var action = step.Action.ToLowerInvariant();
        var parts = new List<string>();

        // Include the most relevant fields per action type
        switch (action)
        {
            case "find":
                if (step.AutomationId != null) parts.Add($"automation_id={SubstituteParams(step.AutomationId, parameters)}");
                if (step.Name != null) parts.Add($"name={SubstituteParams(step.Name, parameters)}");
                if (step.ClassName != null) parts.Add($"class_name={SubstituteParams(step.ClassName, parameters)}");
                if (step.ControlType != null) parts.Add($"control_type={SubstituteParams(step.ControlType, parameters)}");
                if (step.SaveAs != null) parts.Add($"save_as={step.SaveAs}");
                break;
            case "find_by_path":
                if (step.Path != null) parts.Add($"path=[{step.Path.Count} segments]");
                if (step.SaveAs != null) parts.Add($"save_as={step.SaveAs}");
                break;
            case "click" or "right_click" or "properties" or "get_value":
                if (step.Ref != null) parts.Add($"ref={step.Ref}");
                break;
            case "type":
                if (step.Text != null) parts.Add($"text={SubstituteParams(step.Text, parameters)}");
                break;
            case "set_value":
                if (step.Ref != null) parts.Add($"ref={step.Ref}");
                var val = step.Value ?? step.Text;
                if (val != null) parts.Add($"value={SubstituteParams(val, parameters)}");
                break;
            case "send_keys" or "keys":
                if (step.Keys != null) parts.Add($"keys={SubstituteParams(step.Keys, parameters)}");
                break;
            case "wait":
                if (step.Seconds.HasValue) parts.Add($"seconds={step.Seconds.Value}");
                break;
            case "attach":
                if (step.ProcessName != null) parts.Add($"process={SubstituteParams(step.ProcessName, parameters)}");
                if (step.Pid.HasValue) parts.Add($"pid={step.Pid.Value}");
                break;
            case "launch":
                if (step.ExePath != null) parts.Add($"exe={SubstituteParams(step.ExePath, parameters)}");
                if (step.IfNotRunning == true) parts.Add("if_not_running");
                break;
            case "wait_for_window":
                if (step.TitleContains != null) parts.Add($"title_contains={SubstituteParams(step.TitleContains, parameters)}");
                break;
            case "wait_for_enabled":
                if (step.Ref != null) parts.Add($"ref={step.Ref}");
                if (step.AutomationId != null) parts.Add($"automation_id={SubstituteParams(step.AutomationId, parameters)}");
                break;
            case "verify":
                if (step.Ref != null) parts.Add($"ref={step.Ref}");
                if (step.Property != null) parts.Add($"property={step.Property}");
                if (step.Expected != null) parts.Add($"expected={SubstituteParams(step.Expected, parameters)}");
                if (step.MatchMode != null) parts.Add($"match_mode={step.MatchMode}");
                break;
            case "snapshot":
                if (step.MaxDepth.HasValue) parts.Add($"depth={step.MaxDepth.Value}");
                break;
            case "file_dialog":
                if (step.Text != null) parts.Add($"path={SubstituteParams(step.Text, parameters)}");
                break;
            case "children":
                if (step.Ref != null) parts.Add($"ref={step.Ref}");
                break;
            case "macro":
            case "include":
                if (step.MacroName != null) parts.Add($"macro={step.MacroName}");
                break;
        }

        return parts.Count > 0 ? $"{action} ({string.Join(", ", parts)})" : action;
    }

    private async Task<MacroResult> ExecuteInternalAsync(
        MacroDefinition macro,
        string macroName,
        Dictionary<string, string>? parameters = null,
        UiaEngine? engine = null,
        ElementCache? cache = null,
        CancellationToken cancellation = default,
        Action<string>? onLog = null)
    {
        engine ??= UiaEngine.Instance;
        cache ??= new ElementCache();
        parameters ??= new Dictionary<string, string>();

        // Validate required parameters
        var validationError = ValidateParameters(macro, parameters);
        if (validationError != null)
            return new MacroResult(false, 0, macro.Steps.Count, validationError);

        // Apply defaults for missing optional parameters
        foreach (var p in macro.Parameters)
        {
            if (!parameters.ContainsKey(p.Name) && p.Default != null)
                parameters[p.Name] = p.Default;
        }

        // Macro-level timeout
        var macroTimeoutSec = macro.Timeout > 0 ? macro.Timeout : Constants.DefaultMacroTimeoutSec;
        using var macroCts = CancellationTokenSource.CreateLinkedTokenSource(cancellation);
        macroCts.CancelAfter(TimeSpan.FromSeconds(macroTimeoutSec));

        // Local aliases for save_as references within this macro run
        var aliases = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

        for (int i = 0; i < macro.Steps.Count; i++)
        {
            var step = macro.Steps[i];

            // Check for cancellation / macro timeout
            if (macroCts.IsCancellationRequested)
                return new MacroResult(false, i, macro.Steps.Count,
                    $"Macro '{macroName}' timed out after {macroTimeoutSec}s at step {i + 1} ({step.Action})",
                    i, step.Action, "Macro timeout exceeded");

            // Check process is still alive (except for actions that don't require attachment)
            if (step.Action is not ("attach" or "wait" or "macro" or "include" or "launch" or "wait_for_window"))
            {
                if (!engine.IsAttached)
                    return new MacroResult(false, i, macro.Steps.Count,
                        $"Target process exited during macro execution at step {i + 1} ({step.Action})",
                        i, step.Action, "Process is no longer attached");
            }

            var stepSummary = FormatStepSummary(step, parameters);
            onLog?.Invoke($"[Macro] Step {i + 1}/{macro.Steps.Count}: {stepSummary}");

            try
            {
                var stepResult = await ExecuteStepAsync(
                    step, i, parameters, aliases, engine, cache, macroCts.Token, onLog);

                if (!stepResult.Success)
                {
                    onLog?.Invoke($"[Macro] Step {i + 1}/{macro.Steps.Count}: FAILED \u2014 {stepResult.Error ?? stepResult.Message}");
                    return new MacroResult(false, i, macro.Steps.Count,
                        stepResult.Message, i, step.Action, stepResult.Error);
                }

                onLog?.Invoke($"[Macro] Step {i + 1}/{macro.Steps.Count}: OK \u2014 {stepResult.Message}");
            }
            catch (OperationCanceledException)
            {
                onLog?.Invoke($"[Macro] Step {i + 1}/{macro.Steps.Count}: TIMEOUT");
                return new MacroResult(false, i, macro.Steps.Count,
                    $"Macro '{macroName}' timed out at step {i + 1} ({step.Action})",
                    i, step.Action, "Timeout");
            }
            catch (Exception ex)
            {
                onLog?.Invoke($"[Macro] Step {i + 1}/{macro.Steps.Count}: ERROR \u2014 {ex.Message}");
                return new MacroResult(false, i, macro.Steps.Count,
                    $"Step {i + 1} ({step.Action}) failed: {ex.Message}",
                    i, step.Action, ex.Message);
            }
        }

        onLog?.Invoke($"[Macro] '{macroName}' completed ({macro.Steps.Count} steps)");
        return new MacroResult(true, macro.Steps.Count, macro.Steps.Count,
            $"Macro '{macroName}' completed ({macro.Steps.Count} steps)");
    }

    private async Task<StepResult> ExecuteStepAsync(
        MacroStep step, int index,
        Dictionary<string, string> parameters,
        Dictionary<string, string> aliases,
        UiaEngine engine, ElementCache cache,
        CancellationToken cancellation,
        Action<string>? onLog = null)
    {
        var stepTimeoutSec = step.StepTimeout ?? Constants.DefaultStepTimeoutSec;
        using var stepCts = CancellationTokenSource.CreateLinkedTokenSource(cancellation);
        stepCts.CancelAfter(TimeSpan.FromSeconds(stepTimeoutSec));

        switch (step.Action.ToLowerInvariant())
        {
            case "focus":
            {
                var r = engine.FocusWindow();
                return new StepResult(r.success, r.message);
            }

            case "attach":
            {
                var procName = SubstituteParams(step.ProcessName, parameters);
                var r = step.Pid.HasValue
                    ? engine.AttachByPid(step.Pid.Value)
                    : engine.Attach(procName ?? "");
                return new StepResult(r.success, r.message);
            }

            case "snapshot":
            {
                var depth = step.MaxDepth ?? Constants.DefaultSnapshotDepth;
                var r = engine.GetSnapshot(depth);
                return new StepResult(r.success, r.tree);
            }

            case "find":
            {
                var automationId = SubstituteParams(step.AutomationId, parameters);
                var name = SubstituteParams(step.Name, parameters);
                var className = SubstituteParams(step.ClassName, parameters);
                var controlType = SubstituteParams(step.ControlType, parameters);
                var retryInterval = step.RetryInterval ?? Constants.DefaultRetryIntervalSec;

                // Retry loop for find actions
                while (true)
                {
                    var r = engine.FindElement(automationId, name, className, controlType);
                    if (r.success && r.element != null)
                    {
                        var refKey = cache.Add(r.element);
                        if (step.SaveAs != null)
                            aliases[step.SaveAs] = refKey;
                        return new StepResult(true, $"Found [{refKey}]: {r.message}");
                    }

                    if (stepCts.IsCancellationRequested)
                        return new StepResult(false, $"Element not found after {stepTimeoutSec}s",
                            $"find: automation_id={automationId}, name={name}, class_name={className}, control_type={controlType}");

                    await Task.Delay(TimeSpan.FromSeconds(retryInterval), stepCts.Token);
                }
            }

            case "find_by_path":
            {
                var path = step.Path ?? new List<string>();
                // Substitute params in each path segment
                path = path.Select(s => SubstituteParams(s, parameters) ?? s).ToList();
                var retryInterval = step.RetryInterval ?? Constants.DefaultRetryIntervalSec;

                while (true)
                {
                    var r = engine.FindElementByPath(path);
                    if (r.success && r.element != null)
                    {
                        var refKey = cache.Add(r.element);
                        if (step.SaveAs != null)
                            aliases[step.SaveAs] = refKey;
                        return new StepResult(true, $"Found [{refKey}]: {r.message}");
                    }

                    if (stepCts.IsCancellationRequested)
                        return new StepResult(false, $"Path element not found after {stepTimeoutSec}s",
                            $"find_by_path: {string.Join(" > ", path)}");

                    await Task.Delay(TimeSpan.FromSeconds(retryInterval), stepCts.Token);
                }
            }

            case "children":
            {
                var refKey = ResolveRef(step.Ref, aliases);
                System.Windows.Automation.AutomationElement? parent = null;
                if (refKey != null && !cache.TryGet(refKey, out parent))
                    return new StepResult(false, $"Unknown ref '{refKey}'");
                var children = engine.GetChildElements(parent);
                // Cache all children; if save_as, store first child ref
                string? firstRef = null;
                foreach (var child in children)
                {
                    var r = cache.Add(child);
                    firstRef ??= r;
                }
                if (step.SaveAs != null && firstRef != null)
                    aliases[step.SaveAs] = firstRef;
                return new StepResult(true, $"Found {children.Count} children");
            }

            case "click":
            {
                var refKey = ResolveRef(step.Ref, aliases);
                if (refKey == null)
                    return new StepResult(false, "click requires a ref");
                if (!cache.TryGet(refKey, out var el))
                    return new StepResult(false, $"Unknown ref '{refKey}'");
                var r = engine.ClickElement(el!);
                return new StepResult(r.success, r.message);
            }

            case "right_click":
            {
                var refKey = ResolveRef(step.Ref, aliases);
                if (refKey == null)
                    return new StepResult(false, "right_click requires a ref");
                if (!cache.TryGet(refKey, out var el))
                    return new StepResult(false, $"Unknown ref '{refKey}'");
                var r = engine.RightClickElement(el!);
                return new StepResult(r.success, r.message);
            }

            case "type":
            {
                var text = SubstituteParams(step.Text, parameters);
                if (text == null)
                    return new StepResult(false, "type requires text");
                var r = engine.TypeText(text);
                return new StepResult(r.success, r.message);
            }

            case "send_keys":
            case "keys":
            {
                var keys = SubstituteParams(step.Keys, parameters);
                if (keys == null)
                    return new StepResult(false, "send_keys requires keys");
                var r = engine.SendKeyboardShortcut(keys);
                return new StepResult(r.success, r.message);
            }

            case "set_value":
            {
                var refKey = ResolveRef(step.Ref, aliases);
                if (refKey == null)
                    return new StepResult(false, "set_value requires a ref");
                if (!cache.TryGet(refKey, out var el))
                    return new StepResult(false, $"Unknown ref '{refKey}'");
                // Prefer 'value' property; fall back to 'text' for backward compatibility
                var text = SubstituteParams(step.Value ?? step.Text, parameters);
                if (text == null)
                    return new StepResult(false, "set_value requires a value (use 'value' or 'text' field)");
                var r = engine.SetElementValue(el!, text);
                return new StepResult(r.success, r.message);
            }

            case "get_value":
            {
                var refKey = ResolveRef(step.Ref, aliases);
                if (refKey == null)
                    return new StepResult(false, "get_value requires a ref");
                if (!cache.TryGet(refKey, out var el))
                    return new StepResult(false, $"Unknown ref '{refKey}'");
                var r = engine.GetElementValue(el!);
                return new StepResult(r.success, r.message);
            }

            case "verify":
            {
                var refKey = ResolveRef(step.Ref, aliases);
                if (refKey == null)
                    return new StepResult(false, "verify requires a ref");
                if (!cache.TryGet(refKey, out var el))
                    return new StepResult(false, $"Unknown ref '{refKey}'");

                var property = SubstituteParams(step.Property, parameters);
                if (string.IsNullOrEmpty(property))
                    return new StepResult(false, "verify requires a property");

                var expected = SubstituteParams(step.Expected, parameters);
                if (expected == null)
                    return new StepResult(false, "verify requires an expected value");

                var readResult = engine.ReadElementProperty(el!, property);
                if (!readResult.success)
                    return new StepResult(false, readResult.message);

                var matchMode = step.MatchMode?.ToLowerInvariant() ?? "equals";
                var actual = readResult.value ?? "";
                var matched = VerifyMatch(actual, expected, matchMode);

                if (matched == null)
                    return new StepResult(false, $"Unknown match_mode '{step.MatchMode}'. Valid: equals, contains, not_equals, regex, starts_with");

                if (matched.Value)
                    return new StepResult(true, $"Verify passed ({matchMode}): {property} = \"{actual}\"");

                var failMsg = step.Message != null
                    ? SubstituteParams(step.Message, parameters)
                    : $"Verify failed ({matchMode}): expected {property} {MatchModeDescription(matchMode)} \"{expected}\" but got \"{actual}\"";
                return new StepResult(false, failMsg ?? "Verify failed");
            }

            case "file_dialog":
            {
                var filePath = SubstituteParams(step.Text, parameters);
                if (string.IsNullOrEmpty(filePath))
                    return new StepResult(false, "file_dialog requires text (the file path)");
                var r = engine.FileDialogSetPath(filePath);
                return new StepResult(r.success, r.message);
            }

            case "screenshot":
            {
                var r = engine.TakeScreenshot();
                return new StepResult(r.success, r.message);
            }

            case "properties":
            {
                var refKey = ResolveRef(step.Ref, aliases);
                if (refKey == null)
                    return new StepResult(false, "properties requires a ref");
                if (!cache.TryGet(refKey, out var el))
                    return new StepResult(false, $"Unknown ref '{refKey}'");
                var propsDict = engine.GetElementProperties(el!);
                var propsSummary = string.Join(", ", propsDict.Select(kv => $"{kv.Key}={kv.Value}"));
                return new StepResult(true, $"Properties: {propsSummary}");
            }

            case "wait":
            {
                var seconds = step.Seconds ?? 1;
                await Task.Delay(TimeSpan.FromSeconds(seconds), stepCts.Token);
                return new StepResult(true, $"Waited {seconds}s");
            }

            case "launch":
            {
                var exePath = SubstituteParams(step.ExePath, parameters);
                if (string.IsNullOrEmpty(exePath))
                    return new StepResult(false, "launch requires exe_path");

                var arguments = SubstituteParams(step.Arguments, parameters);
                var workingDir = SubstituteParams(step.WorkingDirectory, parameters);
                var ifNotRunning = step.IfNotRunning ?? true;
                var launchTimeout = step.StepTimeout ?? Constants.DefaultLaunchTimeoutSec;

                var r = await engine.LaunchAndAttachAsync(
                    exePath, arguments, workingDir, ifNotRunning, launchTimeout, stepCts.Token);
                return new StepResult(r.success, r.message);
            }

            case "wait_for_window":
            {
                var titleContains = SubstituteParams(step.TitleContains, parameters);
                var automationId = SubstituteParams(step.AutomationId, parameters);
                var name = SubstituteParams(step.Name, parameters);
                var controlType = SubstituteParams(step.ControlType, parameters);
                var waitTimeout = step.StepTimeout ?? Constants.DefaultLaunchTimeoutSec;
                var pollMs = (int)((step.RetryInterval ?? Constants.DefaultRetryIntervalSec) * 1000);

                var r = await engine.WaitForWindowReadyAsync(
                    titleContains, automationId, name, controlType,
                    waitTimeout, pollMs, stepCts.Token);
                return new StepResult(r.success, r.message);
            }

            case "wait_for_enabled":
            {
                var targetEnabled = step.Enabled ?? true;
                var retryInterval = step.RetryInterval ?? Constants.DefaultRetryIntervalSec;

                // If a ref is provided, use the cached element directly
                var refKey = ResolveRef(step.Ref, aliases);
                if (refKey != null)
                {
                    if (!cache.TryGet(refKey, out var el))
                        return new StepResult(false, $"Unknown ref '{refKey}'");

                    while (true)
                    {
                        var isEnabled = engine.IsElementEnabled(el!);
                        if (isEnabled == targetEnabled)
                            return new StepResult(true,
                                $"Element [{refKey}] IsEnabled={isEnabled} (target={targetEnabled})");

                        if (stepCts.IsCancellationRequested)
                            return new StepResult(false,
                                $"Element [{refKey}] IsEnabled={isEnabled} after {stepTimeoutSec}s (target={targetEnabled})",
                                $"wait_for_enabled: ref={refKey}, target={targetEnabled}");

                        await Task.Delay(TimeSpan.FromSeconds(retryInterval), stepCts.Token);
                    }
                }

                // Otherwise, find by criteria and check IsEnabled
                var automationId = SubstituteParams(step.AutomationId, parameters);
                var name = SubstituteParams(step.Name, parameters);
                var className = SubstituteParams(step.ClassName, parameters);
                var controlType = SubstituteParams(step.ControlType, parameters);

                while (true)
                {
                    var r = engine.FindElement(automationId, name, className, controlType);
                    if (r.success && r.element != null)
                    {
                        var isEnabled = engine.IsElementEnabled(r.element);
                        if (isEnabled == targetEnabled)
                        {
                            var rk = cache.Add(r.element);
                            if (step.SaveAs != null)
                                aliases[step.SaveAs] = rk;
                            return new StepResult(true,
                                $"Element [{rk}] IsEnabled={isEnabled} (target={targetEnabled})");
                        }
                    }

                    if (stepCts.IsCancellationRequested)
                        return new StepResult(false,
                            $"Element not enabled={targetEnabled} after {stepTimeoutSec}s",
                            $"wait_for_enabled: automation_id={automationId}, name={name}, target={targetEnabled}");

                    await Task.Delay(TimeSpan.FromSeconds(retryInterval), stepCts.Token);
                }
            }

            case "macro":
            {
                var nestedName = SubstituteParams(step.MacroName, parameters);
                if (string.IsNullOrEmpty(nestedName))
                    return new StepResult(false, "macro action requires macro_name");

                // Merge nested params with parameter substitution
                var nestedParams = new Dictionary<string, string>();
                if (step.Params != null)
                {
                    foreach (var (k, v) in step.Params)
                        nestedParams[k] = SubstituteParams(v, parameters) ?? v;
                }

                // Use the nested macro's own timeout instead of the default 5s step timeout.
                // Look up the nested macro to read its timeout field; fall back to DefaultMacroTimeoutSec.
                var nestedMacro = Get(nestedName);
                if (nestedMacro == null)
                    return new StepResult(false, $"Macro '{nestedName}' not found");

                var nestedTimeoutSec = step.StepTimeout
                    ?? (nestedMacro.Timeout > 0 ? nestedMacro.Timeout : Constants.DefaultMacroTimeoutSec);
                using var nestedCts = CancellationTokenSource.CreateLinkedTokenSource(cancellation);
                nestedCts.CancelAfter(TimeSpan.FromSeconds(nestedTimeoutSec));

                // Prefix nested macro log messages with parent step context
                var parentPrefix = $"Step {index + 1}";
                Action<string>? nestedLog = onLog != null
                    ? msg => onLog(msg.Replace("[Macro] ", $"[Macro] {parentPrefix} > "))
                    : null;

                var nestedResult = await ExecuteInternalAsync(
                    nestedMacro, nestedName, nestedParams, engine, cache, nestedCts.Token, nestedLog);
                return new StepResult(nestedResult.Success, nestedResult.Message, nestedResult.Error);
            }

            case "include":
                // Include steps should be expanded at load time by ExpandIncludes().
                // If we reach here, the include was not expanded (e.g., macro was
                // constructed at runtime via executeMacroYaml without going through Reload).
                return new StepResult(false,
                    $"include step was not expanded at load time (macro_name={step.MacroName}). " +
                    "Include steps only work in macros loaded from the macros/ folder.");

            default:
                return new StepResult(false, $"Unknown action: {step.Action}");
        }
    }

    /// <summary>Resolve a ref that might be an alias from save_as.</summary>
    private static string? ResolveRef(string? refOrAlias, Dictionary<string, string> aliases)
    {
        if (refOrAlias == null) return null;
        return aliases.TryGetValue(refOrAlias, out var resolved) ? resolved : refOrAlias;
    }

    /// <summary>Replace {{paramName}} placeholders with parameter values.</summary>
    public static string? SubstituteParams(string? template, Dictionary<string, string> parameters)
    {
        if (template == null) return null;
        return s_paramPattern.Replace(template, match =>
        {
            var paramName = match.Groups[1].Value;
            return parameters.TryGetValue(paramName, out var val) ? val : match.Value;
        });
    }

    /// <summary>Validate that all required parameters are provided.</summary>
    public static string? ValidateParameters(MacroDefinition macro, Dictionary<string, string> parameters)
    {
        var missing = macro.Parameters
            .Where(p => p.Required && !parameters.ContainsKey(p.Name) && p.Default == null)
            .Select(p => p.Name)
            .ToList();

        if (missing.Count > 0)
            return $"Missing required parameters: {string.Join(", ", missing)}";

        return null;
    }

    /// <summary>Internal result for a single step execution.</summary>
    internal record StepResult(bool Success, string Message, string? Error = null);

    /// <summary>
    /// Evaluate a verify match. Returns true/false on match result, or null for unknown match mode.
    /// </summary>
    internal static bool? VerifyMatch(string actual, string expected, string matchMode)
    {
        return matchMode switch
        {
            "equals" => string.Equals(actual, expected, StringComparison.OrdinalIgnoreCase),
            "contains" => actual.Contains(expected, StringComparison.OrdinalIgnoreCase),
            "not_equals" => !string.Equals(actual, expected, StringComparison.OrdinalIgnoreCase),
            "regex" => Regex.IsMatch(actual, expected, RegexOptions.IgnoreCase),
            "starts_with" => actual.StartsWith(expected, StringComparison.OrdinalIgnoreCase),
            _ => null
        };
    }

    /// <summary>
    /// Human-readable description of a match mode for failure messages.
    /// </summary>
    private static string MatchModeDescription(string matchMode)
    {
        return matchMode switch
        {
            "equals" => "=",
            "contains" => "to contain",
            "not_equals" => "!=",
            "regex" => "to match pattern",
            "starts_with" => "to start with",
            _ => "="
        };
    }
}

/// <summary>Macro metadata returned by the list command.</summary>
public record MacroInfo(
    string Name,
    string DisplayName,
    string Description,
    List<MacroParamInfo> Parameters);

public record MacroParamInfo(
    string Name,
    string Description,
    bool Required,
    string? Default);

/// <summary>Describes a YAML file that failed to load.</summary>
public record MacroLoadError(
    string FilePath,
    string MacroName,
    string Error);
