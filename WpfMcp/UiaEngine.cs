using System.Diagnostics;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Automation;
using System.Windows.Input;

namespace WpfMcp;

/// <summary>
/// Core UI Automation engine using managed System.Windows.Automation.
/// All UIA operations are dispatched to a dedicated STA thread, which is
/// required for proper COM marshaling to WPF automation peers.
/// Without STA, the UIA COM client only sees the Win32 window layer
/// (TitleBar + FrameworkId="Win32") and cannot traverse into WPF content.
/// </summary>
public class UiaEngine
{
    private static readonly Lazy<UiaEngine> _instance = new(() => new UiaEngine());
    public static UiaEngine Instance => _instance.Value;

    private Process? _attachedProcess;
    private AutomationElement? _mainWindow;

    // Dedicated STA thread for all UIA operations
    private readonly Thread _staThread;
    private readonly BlockingQueue _taskQueue = new();

    public bool IsAttached => _attachedProcess != null && _mainWindow != null && !_attachedProcess.HasExited;
    public string? WindowTitle { get { try { return RunOnSta(() => _mainWindow?.Current.Name); } catch { return null; } } }
    public int? ProcessId => _attachedProcess?.Id;
    public string? ProcessName { get { try { return _attachedProcess?.ProcessName; } catch { return null; } } }
    public IntPtr TargetWindowHandle => _attachedProcess?.MainWindowHandle ?? IntPtr.Zero;

    [DllImport("user32.dll")]
    private static extern bool SetForegroundWindow(IntPtr hWnd);
    [DllImport("user32.dll")]
    private static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);
    [DllImport("user32.dll")]
    private static extern void mouse_event(uint dwFlags, int dx, int dy, uint dwData, UIntPtr dwExtraInfo);
    [DllImport("user32.dll")]
    private static extern bool SetCursorPos(int X, int Y);
    [DllImport("user32.dll")]
    private static extern void keybd_event(byte bVk, byte bScan, uint dwFlags, UIntPtr dwExtraInfo);
    [DllImport("user32.dll")]
    private static extern bool GetWindowRect(IntPtr hWnd, out RECT lpRect);
    [DllImport("user32.dll")]
    private static extern short VkKeyScan(char ch);
    [DllImport("user32.dll")]
    private static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint lpdwProcessId);
    [DllImport("kernel32.dll")]
    private static extern uint GetCurrentThreadId();
    [DllImport("user32.dll")]
    private static extern bool AttachThreadInput(uint idAttach, uint idAttachTo, bool fAttach);
    [DllImport("user32.dll")]
    private static extern IntPtr GetForegroundWindow();
    [DllImport("user32.dll")]
    private static extern bool BringWindowToTop(IntPtr hWnd);
    [DllImport("user32.dll", SetLastError = true)]
    private static extern bool SystemParametersInfo(uint uiAction, uint uiParam, IntPtr pvParam, uint fWinIni);
    private const uint SPI_SETFOREGROUNDLOCKTIMEOUT = 0x2001;
    [DllImport("user32.dll")]
    private static extern bool GetCursorPos(out POINT lpPoint);
    [StructLayout(LayoutKind.Sequential)]
    private struct POINT { public int X, Y; }
    [DllImport("user32.dll", SetLastError = true)]
    private static extern uint SendInput(uint nInputs, INPUT[] pInputs, int cbSize);
    [StructLayout(LayoutKind.Sequential)]
    private struct RECT { public int Left, Top, Right, Bottom; }

    [StructLayout(LayoutKind.Sequential)]
    private struct INPUT
    {
        public uint type;
        public INPUTUNION u;
    }

    [StructLayout(LayoutKind.Explicit)]
    private struct INPUTUNION
    {
        [FieldOffset(0)] public KEYBDINPUT ki;
        [FieldOffset(0)] public MOUSEINPUT mi;
    }

    [StructLayout(LayoutKind.Sequential)]
    private struct KEYBDINPUT
    {
        public ushort wVk;
        public ushort wScan;
        public uint dwFlags;
        public uint time;
        public IntPtr dwExtraInfo;
    }

    [StructLayout(LayoutKind.Sequential)]
    private struct MOUSEINPUT
    {
        public int dx, dy;
        public uint mouseData, dwFlags, time;
        public IntPtr dwExtraInfo;
    }

    // Keyboard hook P/Invoke
    [DllImport("user32.dll", SetLastError = true)]
    private static extern IntPtr SetWindowsHookEx(int idHook, LowLevelKeyboardProc lpfn, IntPtr hMod, uint dwThreadId);
    [DllImport("user32.dll")]
    private static extern bool UnhookWindowsHookEx(IntPtr hhk);
    [DllImport("user32.dll")]
    private static extern IntPtr CallNextHookEx(IntPtr hhk, int nCode, IntPtr wParam, IntPtr lParam);
    [DllImport("kernel32.dll", CharSet = CharSet.Auto)]
    private static extern IntPtr GetModuleHandle(string? lpModuleName);
    [DllImport("kernel32.dll", CharSet = CharSet.Auto, SetLastError = true)]
    private static extern IntPtr LoadLibrary(string lpFileName);
    [DllImport("user32.dll")]
    private static extern int GetMessage(out MSG lpMsg, IntPtr hWnd, uint wMsgFilterMin, uint wMsgFilterMax);
    [DllImport("user32.dll")]
    private static extern bool TranslateMessage(ref MSG lpMsg);
    [DllImport("user32.dll")]
    private static extern IntPtr DispatchMessage(ref MSG lpMsg);
    [DllImport("user32.dll")]
    private static extern bool PostThreadMessage(uint idThread, uint Msg, IntPtr wParam, IntPtr lParam);

    private delegate IntPtr LowLevelKeyboardProc(int nCode, IntPtr wParam, IntPtr lParam);

    [StructLayout(LayoutKind.Sequential)]
    private struct MSG { public IntPtr hwnd; public uint message; public IntPtr wParam, lParam; public uint time; public POINT pt; }

    [StructLayout(LayoutKind.Sequential)]
    private struct KBDLLHOOKSTRUCT { public uint vkCode, scanCode, flags, time; public IntPtr dwExtraInfo; }

    private const int WH_KEYBOARD_LL = 13;
    private const int WM_KEYDOWN = 0x0100;
    private const int WM_KEYUP = 0x0101;
    private const int WM_SYSKEYDOWN = 0x0104;
    private const int WM_SYSKEYUP = 0x0105;
    private const uint WM_QUIT = 0x0012;

    // Keyboard hook state
    private IntPtr _hookHandle;
    private Thread? _hookThread;
    private uint _hookThreadId;
    private TaskCompletionSource? _hookStopped;
    private LowLevelKeyboardProc? _hookProc; // prevent GC of the delegate
    private Action<string, string>? _keypressCallback; // (keyName, keyCombo)
    private readonly HashSet<ushort> _pressedModifiers = new();

    private const uint INPUT_KEYBOARD = 1;
    private const uint KEYEVENTF_KEYUP = 0x0002;

    private const int SW_RESTORE = 9;
    private const uint MOUSEEVENTF_LEFTDOWN = 0x0002;
    private const uint MOUSEEVENTF_LEFTUP = 0x0004;
    private const uint MOUSEEVENTF_RIGHTDOWN = 0x0008;
    private const uint MOUSEEVENTF_RIGHTUP = 0x0010;

    /// <summary>Simple blocking task queue for the STA thread.</summary>
    private class BlockingQueue
    {
        private readonly Queue<Action> _queue = new();
        private readonly object _lock = new();

        public void Enqueue(Action action)
        {
            lock (_lock) { _queue.Enqueue(action); Monitor.Pulse(_lock); }
        }

        public bool TryDequeue(out Action? action, int timeoutMs)
        {
            lock (_lock)
            {
                if (_queue.Count == 0) Monitor.Wait(_lock, timeoutMs);
                if (_queue.Count > 0) { action = _queue.Dequeue(); return true; }
                action = null; return false;
            }
        }
    }

    private UiaEngine()
    {
        _staThread = new Thread(StaThreadLoop)
        {
            Name = "UIA-STA",
            IsBackground = true
        };
        _staThread.SetApartmentState(ApartmentState.STA);
        _staThread.Start();
    }

    private void StaThreadLoop()
    {
        // Initialize COM on this STA thread
        while (true)
        {
            if (_taskQueue.TryDequeue(out var action, 100))
            {
                try { action!(); }
                catch { /* errors are captured via TaskCompletionSource */ }
            }
        }
    }

    /// <summary>
    /// Dispatch a function to the STA thread and wait for the result.
    /// This is the key mechanism that makes UIA work with WPF apps.
    /// </summary>
    private T RunOnSta<T>(Func<T> func)
    {
        if (Thread.CurrentThread == _staThread)
            return func();

        var tcs = new TaskCompletionSource<T>(TaskCreationOptions.RunContinuationsAsynchronously);
        _taskQueue.Enqueue(() =>
        {
            try { tcs.SetResult(func()); }
            catch (Exception ex) { tcs.SetException(ex); }
        });
        return tcs.Task.GetAwaiter().GetResult();
    }

    public (bool success, string message) Attach(string processName)
    {
        return RunOnSta(() =>
        {
            var processes = Process.GetProcessesByName(processName);
            if (processes.Length == 0)
                return (false, $"No process found with name '{processName}'");
            var proc = processes[0];
            if (proc.MainWindowHandle == IntPtr.Zero)
                return (false, $"Process '{processName}' has no main window");
            try
            {
                _attachedProcess = proc;
                _mainWindow = AutomationElement.FromHandle(proc.MainWindowHandle);
                var name = _mainWindow.Current.Name;
                return (true, $"Attached to '{name}' (PID {proc.Id})");
            }
            catch (Exception ex)
            {
                _attachedProcess = null; _mainWindow = null;
                return (false, $"Failed to attach: {ex.Message}");
            }
        });
    }

    public (bool success, string message) AttachByPid(int pid)
    {
        return RunOnSta(() =>
        {
            try
            {
                var proc = Process.GetProcessById(pid);
                if (proc.MainWindowHandle == IntPtr.Zero)
                    return (false, $"Process PID {pid} has no main window");
                _attachedProcess = proc;
                _mainWindow = AutomationElement.FromHandle(proc.MainWindowHandle);
                var name = _mainWindow.Current.Name;
                return (true, $"Attached to '{name}' (PID {proc.Id})");
            }
            catch (Exception ex) { return (false, $"Failed to attach: {ex.Message}"); }
        });
    }

    /// <summary>
    /// Launch a process and attach to it. If <paramref name="ifNotRunning"/> is true
    /// and a process with a matching name is already running, attach to it instead.
    /// Polls until the process has a main window, then calls Attach.
    /// </summary>
    public async Task<(bool success, string message)> LaunchAndAttachAsync(
        string exePath,
        string? arguments = null,
        string? workingDirectory = null,
        bool ifNotRunning = true,
        int timeoutSec = 0,
        CancellationToken cancellation = default)
    {
        if (string.IsNullOrEmpty(exePath))
            return (false, "exe_path is required for launch");

        var exeName = Path.GetFileNameWithoutExtension(exePath);
        timeoutSec = timeoutSec > 0 ? timeoutSec : Constants.DefaultLaunchTimeoutSec;

        // Check if already running
        if (ifNotRunning)
        {
            var existing = Process.GetProcessesByName(exeName);
            if (existing.Length > 0)
            {
                var proc = existing[0];
                if (proc.MainWindowHandle != IntPtr.Zero)
                {
                    var attachResult = AttachByPid(proc.Id);
                    if (attachResult.success)
                        return (true, $"Already running — {attachResult.message}");
                }
                // Process exists but no window yet — fall through to wait for window
                return await WaitForProcessWindowAsync(proc, exeName, timeoutSec, cancellation);
            }
        }

        // Launch the process
        try
        {
            var startInfo = new ProcessStartInfo
            {
                FileName = exePath,
                UseShellExecute = true
            };
            if (!string.IsNullOrEmpty(arguments))
                startInfo.Arguments = arguments;
            if (!string.IsNullOrEmpty(workingDirectory))
                startInfo.WorkingDirectory = workingDirectory;

            var proc = Process.Start(startInfo);
            if (proc == null)
                return (false, $"Failed to start process: {exePath}");

            return await WaitForProcessWindowAsync(proc, exeName, timeoutSec, cancellation);
        }
        catch (Exception ex)
        {
            return (false, $"Failed to launch '{exePath}': {ex.Message}");
        }
    }

    /// <summary>
    /// Poll until a process has a main window, then attach to it.
    /// Handles "relay launch" apps (e.g., Windows 11 Notepad) where the initial
    /// process exits immediately and a new one spawns under a different PID.
    /// When the tracked process exits, falls back to searching by process name.
    /// </summary>
    private async Task<(bool success, string message)> WaitForProcessWindowAsync(
        Process proc, string exeName, int timeoutSec, CancellationToken cancellation)
    {
        using var cts = CancellationTokenSource.CreateLinkedTokenSource(cancellation);
        cts.CancelAfter(TimeSpan.FromSeconds(timeoutSec));

        try
        {
            while (!cts.IsCancellationRequested)
            {
                proc.Refresh();
                if (proc.HasExited)
                {
                    // Process exited — it may have relaunched under a different PID
                    // (common with store apps, e.g., Windows 11 Notepad).
                    // Fall back to searching by process name.
                    return await WaitForProcessByNameAsync(exeName, timeoutSec, cts.Token);
                }

                if (proc.MainWindowHandle != IntPtr.Zero)
                {
                    var attachResult = AttachByPid(proc.Id);
                    if (attachResult.success)
                        return (true, $"Launched and {attachResult.message}");
                    // Window appeared but attach failed — retry briefly
                }

                await Task.Delay(Constants.ProcessMainWindowPollMs, cts.Token);
            }
        }
        catch (OperationCanceledException) { }

        return (false, $"Timed out waiting for '{exeName}' to show a main window ({timeoutSec}s)");
    }

    /// <summary>
    /// Fallback: poll by process name until a matching process with a main window appears.
    /// Used when the initially launched process exits (relay/store apps).
    /// </summary>
    private async Task<(bool success, string message)> WaitForProcessByNameAsync(
        string exeName, int timeoutSec, CancellationToken cancellation)
    {
        try
        {
            while (!cancellation.IsCancellationRequested)
            {
                var candidates = Process.GetProcessesByName(exeName);
                foreach (var p in candidates)
                {
                    try
                    {
                        if (p.MainWindowHandle != IntPtr.Zero)
                        {
                            var attachResult = AttachByPid(p.Id);
                            if (attachResult.success)
                                return (true, $"Launched (relaunched) and {attachResult.message}");
                        }
                    }
                    catch { /* process may have exited between enumeration and access */ }
                }

                await Task.Delay(Constants.ProcessMainWindowPollMs, cancellation);
            }
        }
        catch (OperationCanceledException) { }

        return (false, $"Timed out waiting for '{exeName}' to appear ({timeoutSec}s)");
    }

    /// <summary>
    /// Wait until the attached window is ready. Optionally checks that the window title
    /// contains a substring and/or that a specific element exists in the UIA tree.
    /// </summary>
    public async Task<(bool success, string message)> WaitForWindowReadyAsync(
        string? titleContains = null,
        string? automationId = null,
        string? name = null,
        string? controlType = null,
        int timeoutSec = 0,
        int pollMs = 0,
        CancellationToken cancellation = default)
    {
        timeoutSec = timeoutSec > 0 ? timeoutSec : Constants.DefaultLaunchTimeoutSec;
        pollMs = pollMs > 0 ? pollMs : Constants.DefaultWindowReadyPollMs;

        bool hasElementCriteria = !string.IsNullOrEmpty(automationId)
            || !string.IsNullOrEmpty(name)
            || !string.IsNullOrEmpty(controlType);

        using var cts = CancellationTokenSource.CreateLinkedTokenSource(cancellation);
        cts.CancelAfter(TimeSpan.FromSeconds(timeoutSec));

        try
        {
            while (!cts.IsCancellationRequested)
            {
                if (!IsAttached)
                {
                    await Task.Delay(pollMs, cts.Token);
                    continue;
                }

                // Check title (try UIA first, fall back to Process.MainWindowTitle)
                if (!string.IsNullOrEmpty(titleContains))
                {
                    var title = WindowTitle;
                    if (string.IsNullOrEmpty(title))
                    {
                        // Fallback: read from process directly (works without elevation)
                        try
                        {
                            _attachedProcess?.Refresh();
                            title = _attachedProcess?.MainWindowTitle;
                        }
                        catch { }
                    }
                    if (title == null || !title.Contains(titleContains, StringComparison.OrdinalIgnoreCase))
                    {
                        await Task.Delay(pollMs, cts.Token);
                        continue;
                    }
                }

                // Check element exists
                if (hasElementCriteria)
                {
                    var findResult = FindElement(automationId, name, null, controlType);
                    if (!findResult.success)
                    {
                        await Task.Delay(pollMs, cts.Token);
                        continue;
                    }
                }

                // All criteria met
                var readyMsg = "Window is ready";
                if (!string.IsNullOrEmpty(titleContains))
                    readyMsg += $" (title contains '{titleContains}')";
                if (hasElementCriteria)
                    readyMsg += " (element found)";
                return (true, readyMsg);
            }
        }
        catch (OperationCanceledException) { }

        var failMsg = $"Timed out waiting for window readiness ({timeoutSec}s)";
        if (!IsAttached)
            failMsg += " — not attached to any process";
        else if (!string.IsNullOrEmpty(titleContains))
            failMsg += $" — title '{WindowTitle}' does not contain '{titleContains}'";
        else if (hasElementCriteria)
            failMsg += " — target element not found";
        return (false, failMsg);
    }

    public (bool success, string message) FocusWindow()
    {
        if (!IsAttached) return (false, "Not attached");
        try
        {
            var handle = _attachedProcess!.MainWindowHandle;
            ShowWindow(handle, SW_RESTORE);

            // Temporarily remove the foreground lock timeout so SetForegroundWindow
            // works from a background process. This avoids injecting phantom keystrokes
            // (like Alt) that break macros using send_keys. Works because we're elevated.
            SystemParametersInfo(SPI_SETFOREGROUNDLOCKTIMEOUT, 0, IntPtr.Zero, 0);

            SetForegroundWindow(handle);
            BringWindowToTop(handle);

            Thread.Sleep(Constants.FocusDelayMs);
            return (true, "Window focused");
        }
        catch (Exception ex) { return (false, $"Failed to focus: {ex.Message}"); }
    }

    public (bool success, string tree) GetSnapshot(int maxDepth = Constants.DefaultSnapshotDepth)
    {
        if (!IsAttached) return (false, "Not attached");
        return RunOnSta(() =>
        {
            try
            {
                var sb = new StringBuilder();
                WalkTree(_mainWindow!, sb, 0, maxDepth);
                return (true, sb.ToString());
            }
            catch (Exception ex) { return (false, $"Failed: {ex.Message}"); }
        });
    }

    private void WalkTree(AutomationElement element, StringBuilder sb, int depth, int maxDepth)
    {
        if (depth > maxDepth) return;
        var indent = new string(' ', depth * 2);
        try
        {
            sb.AppendLine($"{indent}{FormatElement(element)}");
        }
        catch { sb.AppendLine($"{indent}[Error reading element]"); return; }

        try
        {
            // Try FindAll with TrueCondition first (more reliable for WPF)
            var children = element.FindAll(TreeScope.Children, Condition.TrueCondition);
            if (children.Count > 0)
            {
                foreach (AutomationElement child in children)
                    WalkTree(child, sb, depth + 1, maxDepth);
            }
            else
            {
                // Fallback to RawViewWalker for elements that don't respond to FindAll
                var walker = TreeWalker.RawViewWalker;
                var child = walker.GetFirstChild(element);
                while (child != null)
                {
                    WalkTree(child, sb, depth + 1, maxDepth);
                    try { child = walker.GetNextSibling(child); } catch { break; }
                }
            }
        }
        catch { }
    }

    public (bool success, AutomationElement? element, string message) FindElementByPath(List<string> pathSegments)
    {
        if (!IsAttached) return (false, null, "Not attached");
        return RunOnSta(() =>
        {
            try
            {
                AutomationElement current = _mainWindow!;
                for (int i = 0; i < pathSegments.Count; i++)
                {
                    var condition = ParsePathSegment(pathSegments[i]);
                    if (condition == null)
                        return (false, (AutomationElement?)null, $"Failed to parse path segment {i}: {pathSegments[i]}");

                    // Try FindFirst on children
                    var found = current.FindFirst(TreeScope.Children, condition);

                    if (found == null)
                    {
                        // Try FindFirst on descendants (deeper search)
                        found = current.FindFirst(TreeScope.Descendants, condition);
                    }

                    if (found == null)
                    {
                        // Fallback: manual walk with RawViewWalker
                        var walker = TreeWalker.RawViewWalker;
                        var child = walker.GetFirstChild(current);
                        while (child != null)
                        {
                            try
                            {
                                if (child.FindFirst(TreeScope.Element, condition) != null)
                                { found = child; break; }
                            }
                            catch { }
                            try { child = walker.GetNextSibling(child); } catch { break; }
                        }
                    }

                    if (found == null)
                        return (false, (AutomationElement?)null, $"Not found at segment {i}: {pathSegments[i]}. Parent: {FormatElement(current)}");
                    current = found;
                }
                return (true, (AutomationElement?)current, FormatElement(current));
            }
            catch (Exception ex) { return (false, (AutomationElement?)null, $"Error: {ex.Message}"); }
        });
    }

    public (bool success, AutomationElement? element, string message) FindElement(
        string? automationId = null, string? name = null,
        string? className = null, string? controlType = null)
    {
        if (!IsAttached) return (false, null, "Not attached");
        return RunOnSta(() =>
        {
            try
            {
                var conditions = new List<Condition>();
                if (!string.IsNullOrEmpty(automationId))
                    conditions.Add(new PropertyCondition(AutomationElement.AutomationIdProperty, automationId));
                if (!string.IsNullOrEmpty(name))
                    conditions.Add(new PropertyCondition(AutomationElement.NameProperty, name));
                if (!string.IsNullOrEmpty(className))
                    conditions.Add(new PropertyCondition(AutomationElement.ClassNameProperty, className));
                if (!string.IsNullOrEmpty(controlType))
                {
                    var ct = GetControlType(controlType);
                    if (ct != null) conditions.Add(new PropertyCondition(AutomationElement.ControlTypeProperty, ct));
                }
                if (conditions.Count == 0) return (false, (AutomationElement?)null, "At least one search property required");

                Condition cond = conditions.Count == 1 ? conditions[0] : new AndCondition(conditions.ToArray());

                // Try FindFirst with Descendants scope
                var found = _mainWindow!.FindFirst(TreeScope.Descendants, cond);

                if (found == null)
                {
                    // Fallback: manual walk with RawViewWalker
                    found = WalkAndFind(_mainWindow!, cond, Constants.MaxWalkDepth);
                }

                if (found == null) return (false, (AutomationElement?)null, "Element not found");
                return (true, (AutomationElement?)found, FormatElement(found));
            }
            catch (Exception ex) { return (false, (AutomationElement?)null, $"Error: {ex.Message}"); }
        });
    }

    private AutomationElement? WalkAndFind(AutomationElement root, Condition condition, int maxDepth, int depth = 0)
    {
        if (depth > maxDepth) return null;
        var walker = TreeWalker.RawViewWalker;
        var child = walker.GetFirstChild(root);
        while (child != null)
        {
            try
            {
                if (child.FindFirst(TreeScope.Element, condition) != null) return child;
                var deeper = WalkAndFind(child, condition, maxDepth, depth + 1);
                if (deeper != null) return deeper;
            }
            catch { }
            try { child = walker.GetNextSibling(child); } catch { break; }
        }
        return null;
    }

    public List<AutomationElement> GetChildElements(AutomationElement? parent = null)
    {
        if (!IsAttached) return new List<AutomationElement>();
        var target = parent ?? _mainWindow!;
        return RunOnSta(() =>
        {
            var result = new List<AutomationElement>();
            try
            {
                // Try FindAll first
                var children = target.FindAll(TreeScope.Children, Condition.TrueCondition);
                if (children.Count > 0)
                {
                    foreach (AutomationElement child in children)
                        result.Add(child);
                    return result;
                }
            }
            catch { }

            // Fallback to RawViewWalker
            try
            {
                var walker = TreeWalker.RawViewWalker;
                var child = walker.GetFirstChild(target);
                while (child != null)
                {
                    result.Add(child);
                    try { child = walker.GetNextSibling(child); } catch { break; }
                }
            }
            catch { }
            return result;
        });
    }

    private (bool success, string message) MouseClickElement(AutomationElement element, uint downFlag, uint upFlag, string verb)
    {
        if (!IsAttached) return (false, "Not attached");
        try
        {
            FocusWindow();
            var rect = RunOnSta(() => element.Current.BoundingRectangle);
            if (rect.IsEmpty) return (false, "No bounding rectangle");
            int x = (int)(rect.Left + rect.Width / 2), y = (int)(rect.Top + rect.Height / 2);
            SetCursorPos(x, y); Thread.Sleep(Constants.PreClickDelayMs);
            mouse_event(downFlag, 0, 0, 0, UIntPtr.Zero); Thread.Sleep(Constants.PreClickDelayMs);
            mouse_event(upFlag, 0, 0, 0, UIntPtr.Zero); Thread.Sleep(Constants.PostClickDelayMs);
            return (true, $"{verb} at ({x}, {y})");
        }
        catch (Exception ex) { return (false, $"{verb} failed: {ex.Message}"); }
    }

    public (bool success, string message) ClickElement(AutomationElement element)
    {
        return MouseClickElement(element, MOUSEEVENTF_LEFTDOWN, MOUSEEVENTF_LEFTUP, "Clicked");
    }

    public (bool success, string message) RightClickElement(AutomationElement element)
    {
        return MouseClickElement(element, MOUSEEVENTF_RIGHTDOWN, MOUSEEVENTF_RIGHTUP, "Right-clicked");
    }

    /// <summary>
    /// Set the value of an element using the UIA ValuePattern.
    /// Works reliably with edit fields, combo boxes, etc. without focus issues.
    /// </summary>
    public (bool success, string message) SetElementValue(AutomationElement element, string value)
    {
        if (!IsAttached) return (false, "Not attached");
        try
        {
            if (element.TryGetCurrentPattern(ValuePattern.Pattern, out object? pattern))
            {
                var valuePattern = (ValuePattern)pattern;
                valuePattern.SetValue(value);
                return (true, $"Set value: {value}");
            }
            return (false, "Element does not support ValuePattern");
        }
        catch (Exception ex) { return (false, $"SetValue failed: {ex.Message}"); }
    }

    /// <summary>
    /// Get the current value of an element using the UIA ValuePattern.
    /// </summary>
    public (bool success, string? value, string message) GetElementValue(AutomationElement element)
    {
        if (!IsAttached) return (false, null, "Not attached");
        try
        {
            if (element.TryGetCurrentPattern(ValuePattern.Pattern, out object? pattern))
            {
                var valuePattern = (ValuePattern)pattern;
                var val = valuePattern.Current.Value;
                return (true, val, $"Value: {val}");
            }
            return (false, null, "Element does not support ValuePattern");
        }
        catch (Exception ex) { return (false, null, $"GetValue failed: {ex.Message}"); }
    }

    /// <summary>
    /// Send keyboard input using SendInput API.
    /// Supports:
    ///   - Simultaneous: "Ctrl+S", "Alt+F4" (modifier held while key pressed)
    ///   - Sequential: "Alt,F" or "Escape" (each key pressed and released independently)
    /// Use comma to separate sequential keypresses, + for simultaneous combos.
    /// </summary>
    public (bool success, string message) SendKeyboardShortcut(string keys)
    {
        if (!IsAttached) return (false, "Not attached");
        try
        {
            // Don't call FocusWindow() — it steals focus from dialogs.
            // Callers should use the explicit 'focus' action when needed.
            Thread.Sleep(Constants.PreKeyDelayMs);

            // Check if this is a sequential key sequence (comma-separated)
            if (keys.Contains(','))
            {
                var sequence = keys.Split(',');
                foreach (var step in sequence)
                {
                    var trimmed = step.Trim();
                    if (string.IsNullOrEmpty(trimmed)) continue;
                    SendKeyCombo(trimmed);
                    Thread.Sleep(Constants.SequentialKeyDelayMs);
                }
                return (true, $"Sent key sequence: {keys}");
            }
            else
            {
                SendKeyCombo(keys);
                return (true, $"Sent keys: {keys}");
            }
        }
        catch (Exception ex) { return (false, $"SendKeys failed: {ex.Message}"); }
    }

    private static readonly HashSet<string> s_modifierNames =
        new(StringComparer.OrdinalIgnoreCase) { "ctrl", "control", "alt", "shift" };

    private void SendKeyCombo(string combo)
    {
        var parts = combo.Split('+');
        var modifiers = new List<ushort>();
        ushort? mainKey = null;

        foreach (var part in parts)
        {
            var trimmed = part.Trim();
            if (s_modifierNames.Contains(trimmed))
                modifiers.Add(ParseVirtualKey(trimmed));
            else
                mainKey = ParseVirtualKey(trimmed);
        }

        var inputs = new List<INPUT>();

        // Press modifiers
        foreach (var mod in modifiers)
            inputs.Add(MakeKeyInput(mod, false));

        // Press and release main key (if any)
        if (mainKey.HasValue)
        {
            inputs.Add(MakeKeyInput(mainKey.Value, false));
            inputs.Add(MakeKeyInput(mainKey.Value, true));
        }

        // Release modifiers (reverse order)
        for (int i = modifiers.Count - 1; i >= 0; i--)
            inputs.Add(MakeKeyInput(modifiers[i], true));

        if (inputs.Count > 0)
        {
            var arr = inputs.ToArray();
            SendInput((uint)arr.Length, arr, Marshal.SizeOf<INPUT>());
        }
    }

    private INPUT MakeKeyInput(ushort vk, bool keyUp)
    {
        return new INPUT
        {
            type = INPUT_KEYBOARD,
            u = new INPUTUNION
            {
                ki = new KEYBDINPUT
                {
                    wVk = vk,
                    wScan = 0,
                    dwFlags = keyUp ? KEYEVENTF_KEYUP : 0,
                    time = 0,
                    dwExtraInfo = IntPtr.Zero
                }
            }
        };
    }

    public (bool success, string message) TypeText(string text)
    {
        if (!IsAttached) return (false, "Not attached");
        try
        {
            // Don't call FocusWindow() — it steals focus from dialogs.
            // Callers should use the explicit 'focus' action when needed.
            Thread.Sleep(Constants.PreTypeDelayMs);
            foreach (char c in text)
            {
                short vk = VkKeyScan(c);
                byte key = (byte)(vk & 0xFF);
                bool shift = (vk & 0x100) != 0;
                if (shift) keybd_event(0x10, 0, 0, UIntPtr.Zero);
                keybd_event(key, 0, 0, UIntPtr.Zero);
                keybd_event(key, 0, KEYEVENTF_KEYUP, UIntPtr.Zero);
                if (shift) keybd_event(0x10, 0, KEYEVENTF_KEYUP, UIntPtr.Zero);
                Thread.Sleep(Constants.PerCharDelayMs);
            }
            return (true, $"Typed: {text}");
        }
        catch (Exception ex) { return (false, $"Type failed: {ex.Message}"); }
    }

    /// <summary>
    /// Fill a standard Windows file dialog with a file path and confirm.
    /// Finds the filename Edit field, clears it, types the full path, and presses Enter.
    /// The file dialog must already be open.
    /// </summary>
    public (bool success, string message) FileDialogSetPath(string filePath)
    {
        if (!IsAttached) return (false, "Not attached");

        // Find the filename Edit field (standard Windows file dialog AutomationId=1148)
        var editResult = FindElement(automationId: "1148", className: "Edit");
        if (!editResult.success || editResult.element == null)
            return (false, "Could not find the filename field (AutomationId=1148). Is a file dialog open?");

        // Click the field to ensure focus
        var clickResult = ClickElement(editResult.element);
        if (!clickResult.success) return (false, $"Failed to click filename field: {clickResult.message}");
        Thread.Sleep(Constants.FileDialogPostClickMs);

        // Clear existing text, type the path, press Enter
        SendKeyboardShortcut("Ctrl+A");
        Thread.Sleep(Constants.FileDialogPostSelectAllMs);
        var typeResult = TypeText(filePath);
        if (!typeResult.success) return (false, $"Failed to type path: {typeResult.message}");
        Thread.Sleep(Constants.FileDialogPreEnterMs);
        SendKeyboardShortcut("Enter");

        return (true, $"File dialog: entered '{filePath}'");
    }

    public (bool success, string base64, string message) TakeScreenshot()
    {
        if (!IsAttached) return (false, "", "Not attached");
        try
        {
            var handle = _attachedProcess!.MainWindowHandle;
            if (!GetWindowRect(handle, out RECT rect)) return (false, "", "Failed to get window rect");
            int w = rect.Right - rect.Left, h = rect.Bottom - rect.Top;
            if (w <= 0 || h <= 0) return (false, "", "Invalid dimensions");
            using var bmp = new System.Drawing.Bitmap(w, h);
            using (var g = System.Drawing.Graphics.FromImage(bmp))
                g.CopyFromScreen(rect.Left, rect.Top, 0, 0, new System.Drawing.Size(w, h));
            using var ms = new System.IO.MemoryStream();
            bmp.Save(ms, System.Drawing.Imaging.ImageFormat.Png);
            return (true, Convert.ToBase64String(ms.ToArray()), $"Screenshot ({w}x{h})");
        }
        catch (Exception ex) { return (false, "", $"Screenshot failed: {ex.Message}"); }
    }

    public Dictionary<string, string> GetElementProperties(AutomationElement element)
    {
        if (!IsAttached) return new Dictionary<string, string>();
        return RunOnSta(() =>
        {
            var props = new Dictionary<string, string>();
            try
            {
                var c = element.Current;
                props["ControlType"] = c.ControlType.ProgrammaticName.Replace("ControlType.", "");
                props["Name"] = c.Name ?? ""; props["AutomationId"] = c.AutomationId ?? "";
                props["ClassName"] = c.ClassName ?? ""; props["IsEnabled"] = c.IsEnabled.ToString();
                props["IsOffscreen"] = c.IsOffscreen.ToString();
                props["BoundingRectangle"] = c.BoundingRectangle.ToString();
                props["FrameworkId"] = c.FrameworkId ?? "";
                props["ProcessId"] = c.ProcessId.ToString();
            }
            catch { }
            try
            {
                var patterns = element.GetSupportedPatterns();
                props["SupportedPatterns"] = string.Join(", ", patterns.Select(p => p.ProgrammaticName.Replace("PatternIdentifiers.Pattern", "")));
            }
            catch { props["SupportedPatterns"] = ""; }

            // Read actual values from supported patterns
            try
            {
                if (element.TryGetCurrentPattern(ValuePattern.Pattern, out var vp))
                    props["Value"] = ((ValuePattern)vp).Current.Value ?? "";
            }
            catch { }
            try
            {
                if (element.TryGetCurrentPattern(TogglePattern.Pattern, out var tp))
                    props["ToggleState"] = ((TogglePattern)tp).Current.ToggleState.ToString();
            }
            catch { }
            try
            {
                if (element.TryGetCurrentPattern(RangeValuePattern.Pattern, out var rvp))
                {
                    var rv = (RangeValuePattern)rvp;
                    props["RangeValue"] = rv.Current.Value.ToString();
                    props["RangeMinimum"] = rv.Current.Minimum.ToString();
                    props["RangeMaximum"] = rv.Current.Maximum.ToString();
                }
            }
            catch { }
            try
            {
                if (element.TryGetCurrentPattern(SelectionItemPattern.Pattern, out var sip))
                    props["IsSelected"] = ((SelectionItemPattern)sip).Current.IsSelected.ToString();
            }
            catch { }
            try
            {
                if (element.TryGetCurrentPattern(ExpandCollapsePattern.Pattern, out var ecp))
                    props["ExpandCollapseState"] = ((ExpandCollapsePattern)ecp).Current.ExpandCollapseState.ToString();
            }
            catch { }
            try
            {
                if (element.TryGetCurrentPattern(SelectionPattern.Pattern, out var sp))
                {
                    var selected = ((SelectionPattern)sp).Current.GetSelection();
                    props["SelectedItems"] = string.Join(", ", selected.Select(s => s.Current.Name ?? ""));
                }
            }
            catch { }
            return props;
        });
    }

    /// <summary>Check whether an element is enabled (runs on STA thread).</summary>
    public bool IsElementEnabled(AutomationElement element)
    {
        return RunOnSta(() => element.Current.IsEnabled);
    }

    /// <summary>
    /// Read a named property from an element. Used by the verify macro step.
    /// Supported properties: value, name, toggle_state, is_enabled, expand_state,
    /// is_selected, control_type, automation_id.
    /// </summary>
    public (bool success, string? value, string message) ReadElementProperty(AutomationElement element, string property)
    {
        if (!IsAttached) return (false, null, "Not attached");
        try
        {
            return RunOnSta<(bool success, string? value, string message)>(() =>
            {
                switch (property.ToLowerInvariant())
                {
                    case "value":
                        if (element.TryGetCurrentPattern(ValuePattern.Pattern, out var vp))
                        {
                            var val = ((ValuePattern)vp).Current.Value ?? "";
                            return (true, val, $"value = \"{val}\"");
                        }
                        return (false, null, "Element does not support ValuePattern");

                    case "name":
                        var name = element.Current.Name ?? "";
                        return (true, name, $"name = \"{name}\"");

                    case "toggle_state":
                        if (element.TryGetCurrentPattern(TogglePattern.Pattern, out var tp))
                        {
                            var state = ((TogglePattern)tp).Current.ToggleState.ToString();
                            return (true, state, $"toggle_state = \"{state}\"");
                        }
                        return (false, null, "Element does not support TogglePattern");

                    case "is_enabled":
                        var enabled = element.Current.IsEnabled.ToString();
                        return (true, enabled, $"is_enabled = \"{enabled}\"");

                    case "expand_state":
                        if (element.TryGetCurrentPattern(ExpandCollapsePattern.Pattern, out var ecp))
                        {
                            var state = ((ExpandCollapsePattern)ecp).Current.ExpandCollapseState.ToString();
                            return (true, state, $"expand_state = \"{state}\"");
                        }
                        return (false, null, "Element does not support ExpandCollapsePattern");

                    case "is_selected":
                        if (element.TryGetCurrentPattern(SelectionItemPattern.Pattern, out var sip))
                        {
                            var selected = ((SelectionItemPattern)sip).Current.IsSelected.ToString();
                            return (true, selected, $"is_selected = \"{selected}\"");
                        }
                        return (false, null, "Element does not support SelectionItemPattern");

                    case "control_type":
                        var ct = element.Current.ControlType.ProgrammaticName.Replace("ControlType.", "");
                        return (true, ct, $"control_type = \"{ct}\"");

                    case "automation_id":
                        var aid = element.Current.AutomationId ?? "";
                        return (true, aid, $"automation_id = \"{aid}\"");

                    default:
                        return (false, null, $"Unknown property '{property}'. Valid: value, name, toggle_state, is_enabled, expand_state, is_selected, control_type, automation_id");
                }
            });
        }
        catch (Exception ex) { return (false, null, $"ReadElementProperty failed: {ex.Message}"); }
    }

    /// <summary>
    /// Returns the currently focused element if it belongs to the attached process.
    /// Returns null if not attached, focus is outside the app, or the element is inaccessible.
    /// </summary>
    public AutomationElement? GetFocusedElement()
    {
        if (!IsAttached) return null;
        try
        {
            return RunOnSta<AutomationElement?>(() =>
            {
                var focused = AutomationElement.FocusedElement;
                if (focused == null) return null;

                // Only return elements belonging to the attached process
                try
                {
                    if (focused.Current.ProcessId != _attachedProcess!.Id)
                        return null;
                }
                catch { return null; }

                return focused;
            });
        }
        catch { return null; }
    }

    /// <summary>
    /// Returns the element under the mouse cursor if it belongs to the attached process.
    /// Uses GetCursorPos + AutomationElement.FromPoint for hit-testing.
    /// Returns null if not attached, cursor is outside the app, or the element is inaccessible.
    /// </summary>
    public AutomationElement? GetElementAtCursor()
    {
        if (!IsAttached) return null;
        try
        {
            return RunOnSta<AutomationElement?>(() =>
            {
                if (!GetCursorPos(out var pt)) return null;
                var element = AutomationElement.FromPoint(new System.Windows.Point(pt.X, pt.Y));
                if (element == null) return null;

                // Only return elements belonging to the attached process
                try
                {
                    if (element.Current.ProcessId != _attachedProcess!.Id)
                        return null;
                }
                catch { return null; }

                return element;
            });
        }
        catch { return null; }
    }

    public string FormatElement(AutomationElement e)
    {
        try
        {
            var c = e.Current;
            return $"[{c.ControlType.ProgrammaticName.Replace("ControlType.", "")}] " +
                   $"Name=\"{c.Name}\" AutomationId=\"{c.AutomationId}\" " +
                   $"ClassName=\"{c.ClassName}\" FrameworkId=\"{c.FrameworkId}\"";
        }
        catch { return "[Error reading element]"; }
    }

    private Condition? ParsePathSegment(string segment)
    {
        var conditions = new List<Condition>();
        foreach (var pair in segment.Split(';', StringSplitOptions.RemoveEmptyEntries))
        {
            var parts = pair.Split('~', 2);
            if (parts.Length != 2) continue;
            var value = parts[1];
            if (value.StartsWith("Contains+")) value = value.Substring("Contains+".Length);
            var colonIdx = parts[0].IndexOf(':');
            if (colonIdx < 0) continue;
            var propName = parts[0].Substring(colonIdx + 1);
            switch (propName)
            {
                case "ControlType":
                    var ct = GetControlType(value);
                    if (ct != null) conditions.Add(new PropertyCondition(AutomationElement.ControlTypeProperty, ct));
                    break;
                case "AutomationId":
                    conditions.Add(new PropertyCondition(AutomationElement.AutomationIdProperty, value)); break;
                case "Name":
                    conditions.Add(new PropertyCondition(AutomationElement.NameProperty, value)); break;
                case "ClassName":
                    conditions.Add(new PropertyCondition(AutomationElement.ClassNameProperty, value)); break;
            }
        }
        if (conditions.Count == 0) return null;
        if (conditions.Count == 1) return conditions[0];
        return new AndCondition(conditions.ToArray());
    }

    /// <summary>
    /// All ControlType static fields, indexed by name (case-insensitive).
    /// Built once via reflection — automatically covers all types without manual maintenance.
    /// </summary>
    private static readonly Dictionary<string, ControlType> s_controlTypes =
        typeof(ControlType)
            .GetFields(BindingFlags.Public | BindingFlags.Static)
            .Where(f => f.FieldType == typeof(ControlType))
            .ToDictionary(f => f.Name, f => (ControlType)f.GetValue(null)!,
                          StringComparer.OrdinalIgnoreCase);

    // ObjectStore aliases not matching ControlType field names
    private static readonly Dictionary<string, string> s_controlTypeAliases =
        new(StringComparer.OrdinalIgnoreCase)
        {
            ["TabList"] = "Tab",
            ["TabPage"] = "TabItem",
        };

    private static ControlType? GetControlType(string name)
    {
        if (s_controlTypeAliases.TryGetValue(name, out var mapped))
            name = mapped;
        return s_controlTypes.TryGetValue(name, out var ct) ? ct : null;
    }

    // --- Keyboard Hook for Watch Mode ---

    /// <summary>
    /// Start a low-level keyboard hook on a dedicated thread with a message pump.
    /// The callback receives (keyName, keyCombo) for each keypress directed at the attached process.
    /// </summary>
    public async Task StartKeyboardHookAsync(Action<string, string> onKeypress)
    {
        await StopKeyboardHookAsync();
        _keypressCallback = onKeypress;
        _pressedModifiers.Clear();

        _hookStopped = new TaskCompletionSource(TaskCreationOptions.RunContinuationsAsynchronously);
        _hookThread = new Thread(KeyboardHookThreadProc);
        _hookThread.SetApartmentState(ApartmentState.STA);
        _hookThread.IsBackground = true;
        _hookThread.Start();
    }

    /// <summary>Stop the keyboard hook and its message pump thread without blocking.</summary>
    public async Task StopKeyboardHookAsync()
    {
        if (_hookThread != null && _hookThreadId != 0)
        {
            PostThreadMessage(_hookThreadId, WM_QUIT, IntPtr.Zero, IntPtr.Zero);
            if (_hookStopped != null)
            {
                // Wait for hook thread to exit, with a safety timeout
                await Task.WhenAny(_hookStopped.Task, Task.Delay(2000));
            }
        }
        _hookThread = null;
        _hookThreadId = 0;
        _hookStopped = null;
        _keypressCallback = null;
        _pressedModifiers.Clear();
    }

    private void KeyboardHookThreadProc()
    {
        _hookThreadId = GetCurrentThreadId();
        _hookProc = HookCallback;

        // Use LoadLibrary("user32.dll") instead of GetModuleHandle(null) —
        // in single-file .NET apps, GetModuleHandle(null) can return a handle
        // that SetWindowsHookEx rejects for WH_KEYBOARD_LL hooks.
        var hMod = LoadLibrary("user32.dll");
        _hookHandle = SetWindowsHookEx(WH_KEYBOARD_LL, _hookProc, hMod, 0);

        if (_hookHandle == IntPtr.Zero)
        {
            var err = Marshal.GetLastWin32Error();
            Console.Error.WriteLine($"[Watch] Failed to install keyboard hook (error {err})");
            _hookStopped?.TrySetResult();
            return;
        }

        Console.Error.WriteLine($"[Watch] Keyboard hook installed (handle={_hookHandle}, thread={_hookThreadId})");

        try
        {
            // Message pump — required for WH_KEYBOARD_LL callbacks to fire
            while (GetMessage(out var msg, IntPtr.Zero, 0, 0) > 0)
            {
                TranslateMessage(ref msg);
                DispatchMessage(ref msg);
            }
        }
        finally
        {
            UnhookWindowsHookEx(_hookHandle);
            _hookHandle = IntPtr.Zero;
            _hookProc = null;
            _hookStopped?.TrySetResult();
        }
    }

    private IntPtr HookCallback(int nCode, IntPtr wParam, IntPtr lParam)
    {
        if (nCode >= 0 && _attachedProcess != null)
        {
            var msg = (uint)wParam;
            var kbd = Marshal.PtrToStructure<KBDLLHOOKSTRUCT>(lParam);
            var vk = (ushort)kbd.vkCode;

            // Track modifier key state (track globally, not just for attached process)
            if (msg == WM_KEYDOWN || msg == WM_SYSKEYDOWN)
            {
                if (IsModifierKey(vk))
                {
                    _pressedModifiers.Add(NormalizeModifier(vk));
                }
                else
                {
                    // Non-modifier key pressed — check if it's directed at the attached process
                    try
                    {
                        var fg = GetForegroundWindow();
                        GetWindowThreadProcessId(fg, out var fgPid);

                        // Also check if the foreground window belongs to a child process
                        // (e.g., file dialogs may be hosted by the same process)
                        var attachedPid = (uint)_attachedProcess.Id;
                        if (fgPid == attachedPid)
                        {
                            var keyName = VkToKeyName(vk);
                            var combo = BuildComboString(keyName);
                            _keypressCallback?.Invoke(keyName, combo);
                        }
                    }
                    catch { /* process may have exited */ }
                }
            }
            else if (msg == WM_KEYUP || msg == WM_SYSKEYUP)
            {
                if (IsModifierKey(vk))
                    _pressedModifiers.Remove(NormalizeModifier(vk));
            }
        }

        return CallNextHookEx(_hookHandle, nCode, wParam, lParam);
    }

    private static bool IsModifierKey(ushort vk) =>
        vk is 0xA0 or 0xA1  // LShift, RShift
            or 0xA2 or 0xA3  // LCtrl, RCtrl
            or 0xA4 or 0xA5  // LAlt, RAlt
            or 0x5B or 0x5C  // LWin, RWin
            or 0x10 or 0x11 or 0x12; // generic Shift, Ctrl, Alt

    /// <summary>Normalize left/right modifier variants to a single VK.</summary>
    private static ushort NormalizeModifier(ushort vk) => vk switch
    {
        0xA0 or 0xA1 or 0x10 => 0x10, // Shift
        0xA2 or 0xA3 or 0x11 => 0x11, // Ctrl
        0xA4 or 0xA5 or 0x12 => 0x12, // Alt
        0x5B or 0x5C => 0x5B,          // Win
        _ => vk
    };

    private string BuildComboString(string keyName)
    {
        var parts = new List<string>();
        if (_pressedModifiers.Contains(0x11)) parts.Add("Ctrl");
        if (_pressedModifiers.Contains(0x12)) parts.Add("Alt");
        if (_pressedModifiers.Contains(0x10)) parts.Add("Shift");
        if (_pressedModifiers.Contains(0x5B)) parts.Add("Win");
        parts.Add(keyName);
        return string.Join("+", parts);
    }

    /// <summary>Reverse-map a virtual key code to a readable key name.</summary>
    private static string VkToKeyName(ushort vk)
    {
        // Try reverse mapping through System.Windows.Input.Key
        try
        {
            var key = KeyInterop.KeyFromVirtualKey(vk);
            if (key != Key.None)
            {
                var name = key.ToString();
                // Clean up D0-D9 → 0-9, Oem names
                if (name.Length == 2 && name[0] == 'D' && char.IsDigit(name[1]))
                    return name[1..];
                return name;
            }
        }
        catch { }
        return $"VK{vk:X2}";
    }

    /// <summary>
    /// Aliases for key names that don't match System.Windows.Input.Key member names.
    /// Bare digits must be mapped because Enum.TryParse interprets "0"-"9" as
    /// raw integer values (yielding wrong keys like Key.Cancel instead of Key.D0).
    /// </summary>
    private static readonly Dictionary<string, string> s_keyAliases =
        new(StringComparer.OrdinalIgnoreCase)
        {
            ["0"] = "D0", ["1"] = "D1", ["2"] = "D2", ["3"] = "D3", ["4"] = "D4",
            ["5"] = "D5", ["6"] = "D6", ["7"] = "D7", ["8"] = "D8", ["9"] = "D9",
            ["Esc"] = "Escape",
            ["Backspace"] = "Back",
            ["Del"] = "Delete",
            ["Ins"] = "Insert",
            ["PgUp"] = "PageUp",
            ["PgDn"] = "PageDown",
            ["Ctrl"] = "LeftCtrl",
            ["Control"] = "LeftCtrl",
            ["Alt"] = "LeftAlt",
            ["Shift"] = "LeftShift",
            // Special characters → OEM key names
            ["/"] = "OemQuestion",
            ["?"] = "OemQuestion",
            [";"] = "OemSemicolon",
            [":"] = "OemSemicolon",
            ["="] = "OemPlus",
            ["-"] = "OemMinus",
            ["."] = "OemPeriod",
            ["["] = "Oem4",
            ["]"] = "Oem6",
            [@"\"] = "Oem5",
            ["`"] = "Oem3",
            ["~"] = "Oem3",
            ["'"] = "Oem7",
        };

    private static ushort ParseVirtualKey(string keyName)
    {
        if (s_keyAliases.TryGetValue(keyName, out var mapped))
            keyName = mapped;
        if (Enum.TryParse<Key>(keyName, ignoreCase: true, out var key) && key != Key.None)
            return (ushort)KeyInterop.VirtualKeyFromKey(key);
        throw new ArgumentException($"Unknown key: '{keyName}'");
    }
}

