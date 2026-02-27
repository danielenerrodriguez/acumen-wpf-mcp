namespace WpfMcp.Web;

/// <summary>
/// Interface for the web dashboard to interact with UIA engine, element cache,
/// and macro engine. Implemented in the main WpfMcp project.
/// Thread-safe via async semaphore to avoid blocking Blazor SignalR threads.
/// </summary>
public interface IAppState
{
    // --- Status ---
    bool IsAttached { get; }
    string? WindowTitle { get; }
    int? ProcessId { get; }

    // --- Server info ---
    string ServerVersion { get; }
    string BuildTime { get; }

    // --- Element tree ---
    Task<List<ElementInfo>> GetRootChildrenAsync();
    Task<List<ElementInfo>> GetChildrenAsync(string refKey);
    Task<Dictionary<string, string>> GetPropertiesAsync(string refKey);
    Task<string> GetSnapshotAsync(int maxDepth = 3);

    // --- Actions ---
    Task<ActionResult> AttachAsync(string processName);
    Task<ActionResult> ClickAsync(string refKey);
    Task<ActionResult> RightClickAsync(string refKey);
    Task<ActionResult> FocusAsync();
    Task<ActionResult> SendKeysAsync(string keys);
    Task<ActionResult> TypeTextAsync(string text);
    Task<ActionResult> SetValueAsync(string refKey, string value);
    Task<ActionResult> GetValueAsync(string refKey);
    Task<ActionResult> VerifyAsync(string refKey, string property, string expected, string matchMode = "equals");

    // --- Find ---
    Task<FindResult?> FindAsync(string? automationId, string? name, string? className, string? controlType);

    // --- Tree navigation ---
    /// <summary>
    /// Get RuntimeId path from root window down to the element identified by refKey.
    /// Returns list of RuntimeId strings in order: [rootChild, ..., parent, target].
    /// Used by the dashboard to auto-expand the tree to a watched element.
    /// </summary>
    Task<List<string>> GetAncestorRuntimeIdsAsync(string refKey);

    // --- Macros ---
    List<MacroSummary> ListMacros();
    Task<MacroRunResult> RunMacroAsync(string name, Dictionary<string, string> parameters);

    /// <summary>Cancel the currently running macro (no-op if nothing is running).</summary>
    void CancelMacro();

    /// <summary>Whether a macro is currently executing.</summary>
    bool IsMacroRunning { get; }

    /// <summary>Open a macro YAML file in the user's default editor (VS Code, Notepad, etc.).</summary>
    ActionResult OpenMacroFile(string macroName);

    /// <summary>
    /// Launch a Windows file browse dialog (OpenFileDialog) and return the selected path.
    /// Returns null if the user cancelled.
    /// </summary>
    Task<string?> BrowseForFileAsync(string? initialPath);

    /// <summary>
    /// Launch a Windows folder browse dialog (FolderBrowserDialog) and return the selected path.
    /// Returns null if the user cancelled.
    /// </summary>
    Task<string?> BrowseForFolderAsync(string? initialPath);

    // --- Processes ---
    /// <summary>List processes that have a visible main window.</summary>
    List<ProcessInfo> ListWindowedProcesses();

    // --- Log ---
    event Action<LogEntry>? OnLog;
    IReadOnlyList<LogEntry> RecentLogs { get; }

    // --- Watch mode ---
    /// <summary>Start a new watch session. Returns the session, or null if already watching.</summary>
    Task<WatchSession?> StartWatchAsync();

    /// <summary>Stop the current watch session. Returns the completed session, or null if not watching.</summary>
    Task<WatchSession?> StopWatchAsync();

    /// <summary>Get the current (active) or last (completed) watch session.</summary>
    WatchSession? GetWatchSession();

    /// <summary>Whether a watch session is currently active.</summary>
    bool IsWatching { get; }

    /// <summary>Fires when a new watch entry is recorded.</summary>
    event Action<WatchEntry>? OnWatchEntry;

    /// <summary>Fires when watch mode starts or stops (from any source: MCP, CLI, or dashboard).</summary>
    event Action? OnWatchStateChanged;

    // --- Events ---
    /// <summary>Fires when attachment status changes (attach/detach).</summary>
    event Action? OnAttachChanged;

    /// <summary>Fires when the macro list is reloaded (FileSystemWatcher detected changes).</summary>
    event Action? OnMacrosChanged;
}

// --- DTOs ---

public record ElementInfo(
    string RefKey,
    string ControlType,
    string Name,
    string AutomationId,
    string ClassName,
    string FrameworkId,
    bool HasChildren,
    string RuntimeId = "");

public record ActionResult(bool Success, string Message);

public record FindResult(ElementInfo Element, Dictionary<string, string> Properties);

public record ProcessInfo(int Pid, string Name, string Title);

// --- Watch session types ---

public enum WatchEntryKind { Focus, Hover, PropertyChange, Keypress }

public record WatchEntry(
    DateTime Time,
    WatchEntryKind Kind,
    string ControlType,
    string AutomationId,
    string Name,
    string? RefKey,
    Dictionary<string, string> Properties,
    string? RuntimeId = null,
    string? ChangedProperty = null,
    string? OldValue = null,
    string? NewValue = null,
    string? KeyName = null,
    string? KeyCombo = null);

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
