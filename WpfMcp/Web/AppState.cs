using System.Windows.Automation;
using WpfMcp.Web;

namespace WpfMcp;

/// <summary>
/// Concrete implementation of <see cref="IAppState"/> that wraps
/// UiaEngine, ElementCache, and MacroEngine for the web dashboard.
/// Thread-safe via the shared command semaphore.
/// </summary>
internal sealed class AppState : IAppState
{
    private readonly UiaEngine _engine;
    private readonly ElementCache _cache;
    private readonly Lazy<MacroEngine> _macroEngine;
    private readonly SemaphoreSlim _lock;
    private readonly List<LogEntry> _recentLogs = new();
    private readonly object _logLock = new();

    public event Action<LogEntry>? OnLog;
    public event Action? OnAttachChanged;
    public event Action<WatchEntry>? OnWatchEntry;

    // Watch session state
    private Timer? _watchTimer;
    private WatchSession? _currentSession;
    private WatchSession? _lastSession;
    private string? _lastFocusedRuntimeId;
    private string? _lastHoverRuntimeId;
    private string? _watchSelectedRef;
    private Dictionary<string, string>? _watchSelectedProps;

    public bool IsWatching => _currentSession?.IsActive == true;

    public AppState(UiaEngine engine, ElementCache cache, Lazy<MacroEngine> macroEngine, SemaphoreSlim commandLock)
    {
        _engine = engine;
        _cache = cache;
        _macroEngine = macroEngine;
        _lock = commandLock;
    }

    // --- Status ---

    public bool IsAttached => _engine.IsAttached;
    public string? WindowTitle => _engine.IsAttached ? _engine.WindowTitle : null;
    public int? ProcessId => _engine.IsAttached ? _engine.ProcessId : null;

    public IReadOnlyList<LogEntry> RecentLogs
    {
        get { lock (_logLock) return _recentLogs.ToList(); }
    }

    // --- Logging ---

    private void Log(Web.LogLevel level, string message, string? refKey = null)
    {
        var entry = new LogEntry(DateTime.Now, level, message, refKey);
        lock (_logLock)
        {
            _recentLogs.Add(entry);
            if (_recentLogs.Count > 500)
                _recentLogs.RemoveRange(0, _recentLogs.Count - 500);
        }
        OnLog?.Invoke(entry);
    }

    // --- Element Tree ---

    public List<ElementInfo> GetRootChildren()
    {
        _lock.Wait();
        try
        {
            var children = _engine.GetChildElements(null);
            return children.Select(el => ToElementInfo(el)).ToList();
        }
        finally { _lock.Release(); }
    }

    public List<ElementInfo> GetChildren(string refKey)
    {
        _lock.Wait();
        try
        {
            if (!_cache.TryGet(refKey, out var parent))
                return new();
            var children = _engine.GetChildElements(parent);
            return children.Select(el => ToElementInfo(el)).ToList();
        }
        finally { _lock.Release(); }
    }

    public Dictionary<string, string> GetProperties(string refKey)
    {
        _lock.Wait();
        try
        {
            if (!_cache.TryGet(refKey, out var el))
                return new();
            return _engine.GetElementProperties(el!);
        }
        finally { _lock.Release(); }
    }

    public string GetSnapshot(int maxDepth = 3)
    {
        _lock.Wait();
        try
        {
            var r = _engine.GetSnapshot(maxDepth);
            return r.tree;
        }
        finally { _lock.Release(); }
    }

    // --- Actions ---

    public ActionResult Attach(string processName)
    {
        _lock.Wait();
        try
        {
            var r = _engine.Attach(processName);
            Log(r.success ? Web.LogLevel.Success : Web.LogLevel.Error, $"Attach: {r.message}");
            if (r.success) OnAttachChanged?.Invoke();
            return new(r.success, r.message);
        }
        finally { _lock.Release(); }
    }

    public ActionResult Click(string refKey)
    {
        _lock.Wait();
        try
        {
            if (!_cache.TryGet(refKey, out var el))
                return LogAndReturn(false, $"Unknown ref '{refKey}'", refKey);
            var r = _engine.ClickElement(el!);
            return LogAndReturn(r.success, $"Click [{refKey}]: {r.message}", refKey);
        }
        finally { _lock.Release(); }
    }

    public ActionResult RightClick(string refKey)
    {
        _lock.Wait();
        try
        {
            if (!_cache.TryGet(refKey, out var el))
                return LogAndReturn(false, $"Unknown ref '{refKey}'", refKey);
            var r = _engine.RightClickElement(el!);
            return LogAndReturn(r.success, $"RightClick [{refKey}]: {r.message}", refKey);
        }
        finally { _lock.Release(); }
    }

    public ActionResult Focus()
    {
        _lock.Wait();
        try
        {
            var r = _engine.FocusWindow();
            return LogAndReturn(r.success, $"Focus: {r.message}");
        }
        finally { _lock.Release(); }
    }

    public ActionResult SendKeys(string keys)
    {
        _lock.Wait();
        try
        {
            var r = _engine.SendKeyboardShortcut(keys);
            return LogAndReturn(r.success, $"Keys '{keys}': {r.message}");
        }
        finally { _lock.Release(); }
    }

    public ActionResult TypeText(string text)
    {
        _lock.Wait();
        try
        {
            var r = _engine.TypeText(text);
            return LogAndReturn(r.success, $"Type: {r.message}");
        }
        finally { _lock.Release(); }
    }

    public ActionResult SetValue(string refKey, string value)
    {
        _lock.Wait();
        try
        {
            if (!_cache.TryGet(refKey, out var el))
                return LogAndReturn(false, $"Unknown ref '{refKey}'", refKey);
            var r = _engine.SetElementValue(el!, value);
            return LogAndReturn(r.success, $"SetValue [{refKey}]: {r.message}", refKey);
        }
        finally { _lock.Release(); }
    }

    public ActionResult GetValue(string refKey)
    {
        _lock.Wait();
        try
        {
            if (!_cache.TryGet(refKey, out var el))
                return LogAndReturn(false, $"Unknown ref '{refKey}'", refKey);
            var r = _engine.GetElementValue(el!);
            return LogAndReturn(r.success, $"GetValue [{refKey}]: {r.message}", refKey);
        }
        finally { _lock.Release(); }
    }

    public ActionResult Verify(string refKey, string property, string expected)
    {
        _lock.Wait();
        try
        {
            if (!_cache.TryGet(refKey, out var el))
                return LogAndReturn(false, $"Unknown ref '{refKey}'");
            var r = _engine.ReadElementProperty(el!, property);
            if (!r.success)
                return LogAndReturn(false, $"Verify [{refKey}]: {r.message}");

            if (string.Equals(r.value, expected, StringComparison.OrdinalIgnoreCase))
            {
                Log(Web.LogLevel.Success, $"Verify PASS: [{refKey}] {property} = \"{r.value}\"", refKey);
                return new(true, $"PASS: {property} = \"{r.value}\"");
            }
            else
            {
                Log(Web.LogLevel.Error, $"Verify FAIL: [{refKey}] expected {property} = \"{expected}\" but got \"{r.value}\"", refKey);
                return new(false, $"FAIL: expected {property} = \"{expected}\" but got \"{r.value}\"");
            }
        }
        finally { _lock.Release(); }
    }

    // --- Find ---

    public FindResult? Find(string? automationId, string? name, string? className, string? controlType)
    {
        _lock.Wait();
        try
        {
            var r = _engine.FindElement(automationId, name, className, controlType);
            if (!r.success || r.element == null)
            {
                Log(Web.LogLevel.Warning, $"Find: {r.message}");
                return null;
            }

            var refKey = _cache.Add(r.element);
            var props = _engine.GetElementProperties(r.element);
            var info = ToElementInfo(r.element, refKey);
            Log(Web.LogLevel.Success, $"Find [{refKey}]: {r.message}", refKey);
            return new FindResult(info, props);
        }
        finally { _lock.Release(); }
    }

    // --- Macros ---

    public List<MacroSummary> ListMacros()
    {
        var macros = _macroEngine.Value.List();
        return macros.Select(m => new MacroSummary(
            m.Name, m.DisplayName, m.Description,
            m.Parameters.Select(p => new MacroParamSummary(p.Name, p.Description, p.Required, p.Default)).ToList()
        )).ToList();
    }

    public async Task<MacroRunResult> RunMacroAsync(string name, Dictionary<string, string> parameters)
    {
        Log(Web.LogLevel.Info, $"Running macro '{name}'...");
        try
        {
            var result = await _macroEngine.Value.ExecuteAsync(name, parameters, _engine, _cache);
            if (result.Success)
                Log(Web.LogLevel.Success, $"Macro '{name}' completed ({result.StepsExecuted}/{result.TotalSteps} steps)");
            else
                Log(Web.LogLevel.Error, $"Macro '{name}' failed: {result.Message}");

            return new MacroRunResult(result.Success, result.Message, result.StepsExecuted, result.TotalSteps, result.Error);
        }
        catch (Exception ex)
        {
            Log(Web.LogLevel.Error, $"Macro '{name}' error: {ex.Message}");
            return new MacroRunResult(false, ex.Message, 0, 0, ex.Message);
        }
    }

    // --- Watch mode (server-side) ---

    public WatchSession? StartWatch()
    {
        if (IsWatching) return null;

        _currentSession = new WatchSession();
        _lastFocusedRuntimeId = null;
        _lastHoverRuntimeId = null;
        _watchSelectedRef = null;
        _watchSelectedProps = null;

        _watchTimer?.Dispose();
        _watchTimer = new Timer(WatchTick, null, 500, 500);

        Log(Web.LogLevel.Info, $"Watch started (session {_currentSession.Id})");
        return _currentSession;
    }

    public WatchSession? StopWatch()
    {
        if (!IsWatching) return _lastSession;

        _watchTimer?.Dispose();
        _watchTimer = null;

        var session = _currentSession!;
        session.StopTime = DateTime.Now;
        _lastSession = session;
        _currentSession = null;

        Log(Web.LogLevel.Info,
            $"Watch stopped (session {session.Id}, {session.Entries.Count} entries, " +
            $"{(session.StopTime.Value - session.StartTime).TotalSeconds:F1}s)");
        return session;
    }

    public WatchSession? GetWatchSession() => _currentSession ?? _lastSession;

    private void WatchTick(object? _)
    {
        if (!IsWatching) return;

        try
        {
            _lock.Wait();
            try
            {
                // --- Focus tracking ---
                bool focusChanged = false;
                ElementInfo? focusInfo = null;
                Dictionary<string, string>? focusProps = null;

                var focusEl = _engine.GetFocusedElement();
                if (focusEl != null)
                {
                    var rid = GetRuntimeId(focusEl);
                    if (!string.IsNullOrEmpty(rid) && rid != _lastFocusedRuntimeId)
                    {
                        _lastFocusedRuntimeId = rid;
                        focusChanged = true;
                        focusInfo = ToElementInfo(focusEl);
                        focusProps = _engine.GetElementProperties(focusEl);

                        var entry = new WatchEntry(
                            DateTime.Now, WatchEntryKind.Focus,
                            focusInfo.ControlType, focusInfo.AutomationId, focusInfo.Name,
                            focusInfo.RefKey, focusProps);
                        RecordWatchEntry(entry);
                        LogWatchEvent("Focus", focusInfo, focusProps);

                        _watchSelectedRef = focusInfo.RefKey;
                        _watchSelectedProps = focusProps;
                    }
                }

                // --- Hover tracking ---
                bool hoverChanged = false;
                var hoverEl = _engine.GetElementAtCursor();
                if (hoverEl != null)
                {
                    var rid = GetRuntimeId(hoverEl);
                    if (!string.IsNullOrEmpty(rid) && rid != _lastHoverRuntimeId)
                    {
                        _lastHoverRuntimeId = rid;
                        hoverChanged = true;
                        var hoverInfo = ToElementInfo(hoverEl);
                        var hoverProps = _engine.GetElementProperties(hoverEl);

                        var entry = new WatchEntry(
                            DateTime.Now, WatchEntryKind.Hover,
                            hoverInfo.ControlType, hoverInfo.AutomationId, hoverInfo.Name,
                            hoverInfo.RefKey, hoverProps);
                        RecordWatchEntry(entry);
                        LogWatchEvent("Hover", hoverInfo, hoverProps);

                        if (!focusChanged)
                        {
                            _watchSelectedRef = hoverInfo.RefKey;
                            _watchSelectedProps = hoverProps;
                        }
                    }
                }

                // --- Property diff on selected element ---
                if (!focusChanged && !hoverChanged && _watchSelectedRef != null && _watchSelectedProps != null)
                {
                    if (_cache.TryGet(_watchSelectedRef, out var selEl) && selEl != null)
                    {
                        var fresh = _engine.GetElementProperties(selEl);
                        if (fresh.Count > 0)
                        {
                            foreach (var kv in fresh)
                            {
                                if (kv.Key == "BoundingRectangle") continue;
                                if (!_watchSelectedProps.TryGetValue(kv.Key, out var old) || old != kv.Value)
                                {
                                    var info = ToElementInfo(selEl, _watchSelectedRef);
                                    var propEntry = new WatchEntry(
                                        DateTime.Now, WatchEntryKind.PropertyChange,
                                        info.ControlType, info.AutomationId, info.Name,
                                        _watchSelectedRef, fresh,
                                        kv.Key, old ?? "(new)", kv.Value);
                                    RecordWatchEntry(propEntry);
                                    Log(Web.LogLevel.Info,
                                        $"{kv.Key}: \"{old ?? "(new)"}\" → \"{kv.Value}\"",
                                        _watchSelectedRef);
                                }
                            }
                            _watchSelectedProps = fresh;
                        }
                    }
                }
            }
            finally { _lock.Release(); }
        }
        catch { /* element may have gone stale */ }
    }

    private void RecordWatchEntry(WatchEntry entry)
    {
        _currentSession?.Entries.Add(entry);
        OnWatchEntry?.Invoke(entry);
    }

    private void LogWatchEvent(string kind, ElementInfo element, Dictionary<string, string> properties)
    {
        var desc = !string.IsNullOrEmpty(element.AutomationId)
            ? $"[{element.ControlType}] #{element.AutomationId}"
            : !string.IsNullOrEmpty(element.Name)
                ? $"[{element.ControlType}] \"{Truncate(element.Name, 30)}\""
                : $"[{element.ControlType}]";

        var extras = new List<string>();
        if (properties.TryGetValue("Value", out var val) && !string.IsNullOrEmpty(val))
            extras.Add($"Value=\"{Truncate(val, 40)}\"");
        if (properties.TryGetValue("ToggleState", out var ts))
            extras.Add($"Toggle={ts}");
        if (properties.TryGetValue("IsSelected", out var sel) && sel == "True")
            extras.Add("Selected");
        if (properties.TryGetValue("ExpandCollapseState", out var ecs) && ecs != "LeafNode")
            extras.Add(ecs);

        var suffix = extras.Count > 0 ? $"  ({string.Join(", ", extras)})" : "";
        Log(Web.LogLevel.Info, $"{kind} → {desc}{suffix}", element.RefKey);
    }

    private static string GetRuntimeId(System.Windows.Automation.AutomationElement el)
    {
        try { return string.Join(".", el.GetRuntimeId()); }
        catch { return ""; }
    }

    private static string Truncate(string s, int max) =>
        s.Length <= max ? s : s[..max] + "...";

    // --- Helpers ---

    private ActionResult LogAndReturn(bool success, string message, string? refKey = null)
    {
        Log(success ? Web.LogLevel.Success : Web.LogLevel.Error, message, refKey);
        return new(success, message);
    }

    private ElementInfo ToElementInfo(AutomationElement el, string? refKey = null)
    {
        var key = refKey ?? _cache.Add(el);
        try
        {
            var c = el.Current;
            var ct = c.ControlType.ProgrammaticName.Replace("ControlType.", "");
            // Check if element has children (quick check via TreeWalker)
            bool hasChildren;
            try
            {
                hasChildren = TreeWalker.ControlViewWalker.GetFirstChild(el) != null;
            }
            catch { hasChildren = false; }

            return new ElementInfo(key, ct, c.Name ?? "", c.AutomationId ?? "", c.ClassName ?? "", c.FrameworkId ?? "", hasChildren);
        }
        catch
        {
            return new ElementInfo(key, "?", "", "", "", "", false);
        }
    }
}
