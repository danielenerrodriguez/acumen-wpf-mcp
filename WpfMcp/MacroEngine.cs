using System.IO;
using System.Text.RegularExpressions;
using YamlDotNet.Serialization;
using YamlDotNet.Serialization.NamingConventions;

namespace WpfMcp;

/// <summary>
/// Loads, validates, and executes macro YAML files.
/// Macros are discovered from the macros/ folder next to the exe,
/// or from the path specified by WPFMCP_MACROS_PATH env var.
/// Watches the macros directory for changes and auto-reloads.
/// </summary>
public class MacroEngine : IDisposable
{
    private readonly Dictionary<string, MacroDefinition> _macros = new(StringComparer.OrdinalIgnoreCase);
    private readonly List<MacroLoadError> _loadErrors = new();
    private readonly string _macrosPath;
    private FileSystemWatcher? _watcher;
    private Timer? _debounceTimer;
    private readonly object _reloadLock = new();
    private bool _disposed;

    private static readonly IDeserializer s_yaml = new DeserializerBuilder()
        .WithNamingConvention(UnderscoredNamingConvention.Instance)
        .IgnoreUnmatchedProperties()
        .Build();

    private static readonly Regex s_paramPattern = new(@"\{\{(\w+)\}\}", RegexOptions.Compiled);

    /// <summary>YAML files that failed to parse on the last reload.</summary>
    public IReadOnlyList<MacroLoadError> LoadErrors
    {
        get { lock (_reloadLock) return _loadErrors.ToList(); }
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
            if (!Directory.Exists(_macrosPath)) return;

            foreach (var file in Directory.GetFiles(_macrosPath, "*.yaml", SearchOption.AllDirectories))
            {
                var relativePath = Path.GetRelativePath(_macrosPath, file);
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

            if (_loadErrors.Count > 0)
                Console.Error.WriteLine($"[{DateTime.Now:HH:mm:ss}] Macros reloaded: {_macros.Count} loaded, {_loadErrors.Count} errors");
            else if (_macros.Count > 0)
                Console.Error.WriteLine($"[{DateTime.Now:HH:mm:ss}] Macros reloaded: {_macros.Count} loaded");
        }
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
        engine ??= UiaEngine.Instance;
        cache ??= new ElementCache();
        parameters ??= new Dictionary<string, string>();

        if (!_macros.TryGetValue(macroName, out var macro))
            return new MacroResult(false, 0, 0, $"Macro '{macroName}' not found");

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

            // Check process is still alive (except for attach/wait/macro actions)
            if (step.Action is not ("attach" or "wait" or "macro"))
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
