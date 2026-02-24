using System.IO;
using System.Text;
using System.Text.Json;
using System.Text.RegularExpressions;
using YamlDotNet.Serialization;
using YamlDotNet.Serialization.NamingConventions;

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
    private readonly string _macrosPath;
    private FileSystemWatcher? _watcher;
    private Timer? _debounceTimer;
    private readonly object _reloadLock = new();
    private bool _disposed;

    private static readonly IDeserializer s_yaml = new DeserializerBuilder()
        .WithNamingConvention(UnderscoredNamingConvention.Instance)
        .IgnoreUnmatchedProperties()
        .Build();

    /// <summary>Deserializer for knowledge base files (generic dictionary).</summary>
    private static readonly IDeserializer s_dictYaml = new DeserializerBuilder()
        .IgnoreUnmatchedProperties()
        .Build();

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
                    var macro = s_yaml.Deserialize<MacroDefinition>(yaml);
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

            var parts = new List<string>();
            if (_macros.Count > 0) parts.Add($"{_macros.Count} macros");
            if (_knowledgeBases.Count > 0) parts.Add($"{_knowledgeBases.Count} knowledge base(s)");
            if (_loadErrors.Count > 0) parts.Add($"{_loadErrors.Count} errors");
            if (parts.Count > 0)
                Console.Error.WriteLine($"[{DateTime.Now:HH:mm:ss}] Reloaded: {string.Join(", ", parts)}");
        }
    }

    /// <summary>Load a _knowledge.yaml file, parse it as a dictionary, and build a summary.</summary>
    private void LoadKnowledgeBase(string filePath, string relativePath)
    {
        try
        {
            var yaml = File.ReadAllText(filePath);
            var dict = s_dictYaml.Deserialize<Dictionary<string, object>>(yaml);
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
        "type", "set_value", "get_value", "wait", "macro", "launch",
        "wait_for_window", "focus", "snapshot", "screenshot", "properties",
        "children", "file_dialog", "attach"
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
                if (!Has("macro_name")) return $"{StepPrefix()}: requires 'macro_name' field";
                break;
            case "wait_for_window":
                if (!Has("title_contains")) return $"{StepPrefix()}: requires 'title_contains' field";
                break;
            case "file_dialog":
                if (!Has("text")) return $"{StepPrefix()}: requires 'text' field (the file path)";
                break;
            case "attach":
                if (!Has("process_name") && !Has("pid"))
                    return $"{StepPrefix()}: requires at least one of: process_name, pid";
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
            foreach (var (folderName, kb) in _knowledgeBases)
            {
                try
                {
                    var dict = s_dictYaml.Deserialize<Dictionary<string, object>>(kb.FullContent);
                    if (dict == null) continue;

                    // Check application.process_name
                    if (dict.TryGetValue("application", out var appObj) && appObj is Dictionary<object, object> app)
                    {
                        if (app.TryGetValue("process_name", out var pn) &&
                            string.Equals(pn?.ToString(), processName, StringComparison.OrdinalIgnoreCase))
                            return folderName;
                    }
                }
                catch { /* skip malformed knowledge bases */ }
            }
        }
        return null;
    }

    /// <summary>
    /// YAML serializer for writing macro files from dictionaries.
    /// Uses underscore naming convention to match the existing YAML format.
    /// </summary>
    private static readonly ISerializer s_yamlSerializer = new SerializerBuilder()
        .WithNamingConvention(UnderscoredNamingConvention.Instance)
        .ConfigureDefaultValuesHandling(DefaultValuesHandling.OmitDefaults)
        .DisableAliases()
        .Build();

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
        var yaml = s_yamlSerializer.Serialize(yamlDoc);

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

    /// <summary>
    /// Execute a macro definition directly (for drag-drop / run-file scenarios).
    /// </summary>
    public async Task<MacroResult> ExecuteDefinitionAsync(
        MacroDefinition macro,
        string displayName,
        Dictionary<string, string>? parameters = null,
        UiaEngine? engine = null,
        ElementCache? cache = null,
        CancellationToken cancellation = default)
    {
        return await ExecuteInternalAsync(macro, displayName, parameters, engine, cache, cancellation);
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
        CancellationToken cancellation = default)
    {
        if (!_macros.TryGetValue(macroName, out var macro))
            return new MacroResult(false, 0, 0, $"Macro '{macroName}' not found");

        return await ExecuteInternalAsync(macro, macroName, parameters, engine, cache, cancellation);
    }

    private async Task<MacroResult> ExecuteInternalAsync(
        MacroDefinition macro,
        string macroName,
        Dictionary<string, string>? parameters = null,
        UiaEngine? engine = null,
        ElementCache? cache = null,
        CancellationToken cancellation = default)
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
            if (step.Action is not ("attach" or "wait" or "macro" or "launch" or "wait_for_window"))
            {
                if (!engine.IsAttached)
                    return new MacroResult(false, i, macro.Steps.Count,
                        $"Target process exited during macro execution at step {i + 1} ({step.Action})",
                        i, step.Action, "Process is no longer attached");
            }

            try
            {
                var stepResult = await ExecuteStepAsync(
                    step, i, parameters, aliases, engine, cache, macroCts.Token);

                if (!stepResult.Success)
                    return new MacroResult(false, i, macro.Steps.Count,
                        stepResult.Message, i, step.Action, stepResult.Error);
            }
            catch (OperationCanceledException)
            {
                return new MacroResult(false, i, macro.Steps.Count,
                    $"Macro '{macroName}' timed out at step {i + 1} ({step.Action})",
                    i, step.Action, "Timeout");
            }
            catch (Exception ex)
            {
                return new MacroResult(false, i, macro.Steps.Count,
                    $"Step {i + 1} ({step.Action}) failed: {ex.Message}",
                    i, step.Action, ex.Message);
            }
        }

        return new MacroResult(true, macro.Steps.Count, macro.Steps.Count,
            $"Macro '{macroName}' completed ({macro.Steps.Count} steps)");
    }

    private async Task<StepResult> ExecuteStepAsync(
        MacroStep step, int index,
        Dictionary<string, string> parameters,
        Dictionary<string, string> aliases,
        UiaEngine engine, ElementCache cache,
        CancellationToken cancellation)
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
                return new StepResult(r.success, r.success ? r.tree : r.tree);
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
                var text = SubstituteParams(step.Text, parameters);
                if (text == null)
                    return new StepResult(false, "set_value requires text");
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
                engine.GetElementProperties(el!);
                return new StepResult(true, "Properties retrieved");
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

                var nestedResult = await ExecuteAsync(nestedName, nestedParams, engine, cache, stepCts.Token);
                return new StepResult(nestedResult.Success, nestedResult.Message, nestedResult.Error);
            }

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
