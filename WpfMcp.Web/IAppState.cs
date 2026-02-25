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

    // --- Watch mode ---
    /// <summary>Start a new watch session. Returns the session, or null if already watching.</summary>
    WatchSession? StartWatch();

    /// <summary>Stop the current watch session. Returns the completed session, or null if not watching.</summary>
    WatchSession? StopWatch();

    /// <summary>Get the current (active) or last (completed) watch session.</summary>
    WatchSession? GetWatchSession();

    /// <summary>Whether a watch session is currently active.</summary>
    bool IsWatching { get; }

    /// <summary>Fires when a new watch entry is recorded.</summary>
    event Action<WatchEntry>? OnWatchEntry;

    // --- Events ---
    /// <summary>Fires when attachment status changes (attach/detach).</summary>
    event Action? OnAttachChanged;
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

// --- Watch session types ---

public enum WatchEntryKind { Focus, Hover, PropertyChange }

public record WatchEntry(
    DateTime Time,
    WatchEntryKind Kind,
    string ControlType,
    string AutomationId,
    string Name,
    string? RefKey,
    Dictionary<string, string> Properties,
    string? ChangedProperty = null,
    string? OldValue = null,
    string? NewValue = null);

public class WatchSession
{
    public string Id { get; } = Guid.NewGuid().ToString("N")[..8];
    public DateTime StartTime { get; } = DateTime.Now;
    public DateTime? StopTime { get; set; }
    public bool IsActive => StopTime == null;
    public List<WatchEntry> Entries { get; } = new();
}

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

public record LogEntry(DateTime Time, LogLevel Level, string Message, string? RefKey = null);

public enum LogLevel { Info, Success, Warning, Error }
