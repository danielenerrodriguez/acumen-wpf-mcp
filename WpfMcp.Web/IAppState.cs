namespace WpfMcp.Web;

/// <summary>
/// Interface for the web dashboard to interact with UIA engine, element cache,
/// and macro engine. Implemented in the main WpfMcp project.
/// </summary>
public interface IAppState
{
    // --- Status ---
    bool IsAttached { get; }
    string? WindowTitle { get; }
    int? ProcessId { get; }

    // --- Element tree ---
    List<ElementInfo> GetRootChildren();
    List<ElementInfo> GetChildren(string refKey);
    Dictionary<string, string> GetProperties(string refKey);
    string GetSnapshot(int maxDepth = 3);

    // --- Actions ---
    ActionResult Attach(string processName);
    ActionResult Click(string refKey);
    ActionResult RightClick(string refKey);
    ActionResult Focus();
    ActionResult SendKeys(string keys);
    ActionResult TypeText(string text);
    ActionResult SetValue(string refKey, string value);
    ActionResult GetValue(string refKey);
    ActionResult Verify(string refKey, string property, string expected);

    // --- Find ---
    FindResult? Find(string? automationId, string? name, string? className, string? controlType);

    // --- Macros ---
    List<MacroSummary> ListMacros();
    Task<MacroRunResult> RunMacroAsync(string name, Dictionary<string, string> parameters);

    // --- Log ---
    event Action<LogEntry>? OnLog;
    IReadOnlyList<LogEntry> RecentLogs { get; }
}

// --- DTOs ---

public record ElementInfo(
    string RefKey,
    string ControlType,
    string Name,
    string AutomationId,
    string ClassName,
    string FrameworkId,
    bool HasChildren);

public record ActionResult(bool Success, string Message);

public record FindResult(ElementInfo Element, Dictionary<string, string> Properties);

public record MacroSummary(
    string Name,
    string DisplayName,
    string Description,
    List<MacroParamSummary> Parameters);

public record MacroParamSummary(
    string Name,
    string Description,
    bool Required,
    string? Default);

public record MacroRunResult(
    bool Success,
    string Message,
    int StepsExecuted,
    int TotalSteps,
    string? Error = null);

public record LogEntry(DateTime Time, LogLevel Level, string Message);

public enum LogLevel { Info, Success, Warning, Error }
