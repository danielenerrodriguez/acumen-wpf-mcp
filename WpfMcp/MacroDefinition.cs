using YamlDotNet.Serialization;

namespace WpfMcp;

/// <summary>
/// YAML-deserialized macro definition. Represents a reusable sequence of
/// UI automation steps with optional parameters.
/// </summary>
public class MacroDefinition
{
    [YamlMember(Alias = "name")]
    public string Name { get; set; } = "";

    [YamlMember(Alias = "description")]
    public string Description { get; set; } = "";

    /// <summary>Max seconds for the entire macro execution. 0 = use default.</summary>
    [YamlMember(Alias = "timeout")]
    public int Timeout { get; set; }

    [YamlMember(Alias = "parameters")]
    public List<MacroParameter> Parameters { get; set; } = new();

    [YamlMember(Alias = "steps")]
    public List<MacroStep> Steps { get; set; } = new();
}

public class MacroParameter
{
    [YamlMember(Alias = "name")]
    public string Name { get; set; } = "";

    [YamlMember(Alias = "description")]
    public string Description { get; set; } = "";

    [YamlMember(Alias = "required")]
    public bool Required { get; set; }

    [YamlMember(Alias = "default")]
    public string? Default { get; set; }
}

/// <summary>
/// A single step in a macro. The <see cref="Action"/> field determines which
/// UIA operation to perform; the remaining properties are action-specific.
/// </summary>
public class MacroStep
{
    [YamlMember(Alias = "action")]
    public string Action { get; set; } = "";

    // --- find / find_by_path ---
    [YamlMember(Alias = "automation_id")]
    public string? AutomationId { get; set; }

    [YamlMember(Alias = "name")]
    public string? Name { get; set; }

    [YamlMember(Alias = "class_name")]
    public string? ClassName { get; set; }

    [YamlMember(Alias = "control_type")]
    public string? ControlType { get; set; }

    [YamlMember(Alias = "path")]
    public List<string>? Path { get; set; }

    /// <summary>Save the found element under this alias for later steps.</summary>
    [YamlMember(Alias = "save_as")]
    public string? SaveAs { get; set; }

    // --- click / right_click / properties / children ---
    [YamlMember(Alias = "ref")]
    public string? Ref { get; set; }

    // --- type ---
    [YamlMember(Alias = "text")]
    public string? Text { get; set; }

    // --- send_keys ---
    [YamlMember(Alias = "keys")]
    public string? Keys { get; set; }

    // --- snapshot ---
    [YamlMember(Alias = "max_depth")]
    public int? MaxDepth { get; set; }

    // --- attach ---
    [YamlMember(Alias = "process_name")]
    public string? ProcessName { get; set; }

    [YamlMember(Alias = "pid")]
    public int? Pid { get; set; }

    // --- wait ---
    [YamlMember(Alias = "seconds")]
    public double? Seconds { get; set; }

    // --- nested macro ---
    /// <summary>Name of nested macro to call (for action: macro).</summary>
    [YamlMember(Alias = "macro_name")]
    public string? MacroName { get; set; }

    [YamlMember(Alias = "params")]
    public Dictionary<string, string>? Params { get; set; }

    // --- launch ---
    /// <summary>Path to the executable to launch (for action: launch).</summary>
    [YamlMember(Alias = "exe_path")]
    public string? ExePath { get; set; }

    /// <summary>Command-line arguments for the launched process.</summary>
    [YamlMember(Alias = "arguments")]
    public string? Arguments { get; set; }

    /// <summary>Working directory for the launched process.</summary>
    [YamlMember(Alias = "working_directory")]
    public string? WorkingDirectory { get; set; }

    /// <summary>
    /// If true (default), skip launch when a matching process is already running
    /// and attach to it instead. Makes launch steps idempotent.
    /// </summary>
    [YamlMember(Alias = "if_not_running")]
    public bool? IfNotRunning { get; set; }

    // --- wait_for_window ---
    /// <summary>Wait until the window title contains this substring.</summary>
    [YamlMember(Alias = "title_contains")]
    public string? TitleContains { get; set; }

    // --- timeouts ---
    /// <summary>Per-step timeout in seconds. 0 = use default.</summary>
    [YamlMember(Alias = "timeout")]
    public int? StepTimeout { get; set; }

    /// <summary>Retry interval in seconds for find actions. 0 = use default.</summary>
    [YamlMember(Alias = "retry_interval")]
    public double? RetryInterval { get; set; }
}

/// <summary>Result returned after macro execution.</summary>
public record MacroResult(
    bool Success,
    int StepsExecuted,
    int TotalSteps,
    string Message,
    int? FailedStepIndex = null,
    string? FailedAction = null,
    string? Error = null);
