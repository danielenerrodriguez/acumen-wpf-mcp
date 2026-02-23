using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Text;

namespace WpfMcp;

/// <summary>
/// Records user interactions (mouse clicks, keyboard input) with the attached
/// WPF application using low-level Windows hooks. Produces a list of
/// RecordedAction objects that can be converted to a macro YAML file.
///
/// Hooks run on the elevated server process, which has the necessary
/// privileges for global hooks and UIA access.
/// </summary>
public class InputRecorder : IDisposable
{
    public enum RecorderState { Idle, Recording }

    private RecorderState _state = RecorderState.Idle;
    private string _macroName = "";
    private string _macrosPath = "";
    private readonly UiaEngine _engine;

    // Captured actions
    private readonly List<RecordedAction> _actions = new();
    private DateTime _recordingStartTime;
    private DateTime _lastActionTime;

    // Keyboard state machine
    private readonly HashSet<ushort> _modifiersDown = new();
    private readonly StringBuilder _typingBuffer = new();
    private DateTime _lastTypingTime;
    private Timer? _typingFlushTimer;
    private bool _altReleasedAlone;
    private Timer? _altSequentialTimer;
    private readonly object _lock = new();

    // Hook handles
    private IntPtr _mouseHook;
    private IntPtr _keyboardHook;
    private LowLevelMouseProc? _mouseDelegate;
    private LowLevelKeyboardProc? _keyboardDelegate;

    // Message pump thread for hooks (low-level hooks require a message loop)
    private Thread? _hookThread;
    private volatile bool _hookThreadRunning;
    private volatile uint _hookThreadNativeId;

    public RecorderState State => _state;
    public int ActionCount { get { lock (_lock) return _actions.Count; } }
    public string MacroName => _macroName;
    public TimeSpan Duration => _state == RecorderState.Recording
        ? DateTime.UtcNow - _recordingStartTime
        : TimeSpan.Zero;

    // P/Invoke for hooks
    private delegate IntPtr LowLevelMouseProc(int nCode, IntPtr wParam, IntPtr lParam);
    private delegate IntPtr LowLevelKeyboardProc(int nCode, IntPtr wParam, IntPtr lParam);

    [DllImport("user32.dll", SetLastError = true)]
    private static extern IntPtr SetWindowsHookEx(int idHook, Delegate lpfn, IntPtr hMod, uint dwThreadId);
    [DllImport("user32.dll")]
    private static extern bool UnhookWindowsHookEx(IntPtr hhk);
    [DllImport("user32.dll")]
    private static extern IntPtr CallNextHookEx(IntPtr hhk, int nCode, IntPtr wParam, IntPtr lParam);
    [DllImport("kernel32.dll")]
    private static extern IntPtr GetModuleHandle(string? lpModuleName);
    [DllImport("user32.dll")]
    private static extern IntPtr GetForegroundWindow();
    [DllImport("user32.dll")]
    private static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint lpdwProcessId);
    [DllImport("user32.dll")]
    private static extern bool GetMessage(out MSG lpMsg, IntPtr hWnd, uint wMsgFilterMin, uint wMsgFilterMax);
    [DllImport("user32.dll")]
    private static extern bool TranslateMessage(ref MSG lpMsg);
    [DllImport("user32.dll")]
    private static extern IntPtr DispatchMessage(ref MSG lpMsg);
    [DllImport("user32.dll")]
    private static extern bool PostThreadMessage(uint idThread, uint Msg, IntPtr wParam, IntPtr lParam);
    [DllImport("kernel32.dll")]
    private static extern uint GetCurrentThreadId();

    private const uint WM_QUIT = 0x0012;

    [StructLayout(LayoutKind.Sequential)]
    private struct MSG
    {
        public IntPtr hwnd;
        public uint message;
        public IntPtr wParam;
        public IntPtr lParam;
        public uint time;
        public POINT pt;
    }

    [StructLayout(LayoutKind.Sequential)]
    private struct POINT { public int X, Y; }

    private const int WH_MOUSE_LL = 14;
    private const int WH_KEYBOARD_LL = 13;
    private const int WM_LBUTTONDOWN = 0x0201;
    private const int WM_RBUTTONDOWN = 0x0204;
    private const int WM_KEYDOWN = 0x0100;
    private const int WM_KEYUP = 0x0101;
    private const int WM_SYSKEYDOWN = 0x0104;
    private const int WM_SYSKEYUP = 0x0105;

    // Virtual key codes for modifiers
    private const ushort VK_LCONTROL = 0xA2;
    private const ushort VK_RCONTROL = 0xA3;
    private const ushort VK_LMENU = 0xA4;    // Left Alt
    private const ushort VK_RMENU = 0xA5;    // Right Alt
    private const ushort VK_LSHIFT = 0xA0;
    private const ushort VK_RSHIFT = 0xA1;
    private const ushort VK_LWIN = 0x5B;
    private const ushort VK_RWIN = 0x5C;
    private const ushort VK_CONTROL = 0x11;
    private const ushort VK_MENU = 0x12;     // Alt
    private const ushort VK_SHIFT = 0x10;

    [StructLayout(LayoutKind.Sequential)]
    private struct MSLLHOOKSTRUCT
    {
        public int x, y;
        public uint mouseData, flags, time;
        public IntPtr dwExtraInfo;
    }

    [StructLayout(LayoutKind.Sequential)]
    private struct KBDLLHOOKSTRUCT
    {
        public uint vkCode, scanCode, flags, time;
        public IntPtr dwExtraInfo;
    }

    public InputRecorder(UiaEngine engine)
    {
        _engine = engine;
    }

    public (bool success, string message) StartRecording(string macroName, string macrosPath)
    {
        lock (_lock)
        {
            if (_state == RecorderState.Recording)
                return (false, "Already recording. Stop the current recording first.");

            if (!_engine.IsAttached)
                return (false, "Not attached to any process. Call wpf_attach first.");

            _macroName = macroName;
            _macrosPath = macrosPath;
            _actions.Clear();
            _modifiersDown.Clear();
            _typingBuffer.Clear();
            _altReleasedAlone = false;
            _recordingStartTime = DateTime.UtcNow;
            _lastActionTime = DateTime.UtcNow;

            // Keep delegates alive as fields to prevent GC collection during hook lifetime
            _mouseDelegate = MouseHookCallback;
            _keyboardDelegate = KeyboardHookCallback;

            // Install hooks on a dedicated thread with a message pump.
            // Low-level hooks (WH_MOUSE_LL, WH_KEYBOARD_LL) require the installing
            // thread to run a message loop — hook callbacks are dispatched via messages.
            var hookReady = new ManualResetEventSlim(false);
            bool hooksInstalled = false;

            _hookThreadRunning = true;
            _hookThread = new Thread(() =>
            {
                _hookThreadNativeId = GetCurrentThreadId();
                var hMod = GetModuleHandle(null);
                _mouseHook = SetWindowsHookEx(WH_MOUSE_LL, _mouseDelegate, hMod, 0);
                _keyboardHook = SetWindowsHookEx(WH_KEYBOARD_LL, _keyboardDelegate, hMod, 0);
                hooksInstalled = _mouseHook != IntPtr.Zero || _keyboardHook != IntPtr.Zero;
                hookReady.Set();

                if (hooksInstalled)
                {
                    // Run message loop to pump hook callbacks
                    while (_hookThreadRunning && GetMessage(out var msg, IntPtr.Zero, 0, 0))
                    {
                        TranslateMessage(ref msg);
                        DispatchMessage(ref msg);
                    }
                }

                // Clean up hooks on this thread
                if (_mouseHook != IntPtr.Zero) { UnhookWindowsHookEx(_mouseHook); _mouseHook = IntPtr.Zero; }
                if (_keyboardHook != IntPtr.Zero) { UnhookWindowsHookEx(_keyboardHook); _keyboardHook = IntPtr.Zero; }
            })
            {
                Name = "InputRecorder-Hooks",
                IsBackground = true
            };
            _hookThread.Start();

            // Wait for hooks to be installed
            hookReady.Wait(TimeSpan.FromSeconds(5));

            if (!hooksInstalled)
            {
                _hookThreadRunning = false;
                _mouseDelegate = null;
                _keyboardDelegate = null;
                return (false, "Failed to install input hooks. Are you running elevated?");
            }

            _state = RecorderState.Recording;
            Console.Error.WriteLine($"[{DateTime.Now:HH:mm:ss}] Recording started: {macroName}");
            return (true, $"Recording started for '{macroName}'. Interact with the application, then stop recording.");
        }
    }

    public (bool success, string message, string? yaml, string? filePath) StopRecording()
    {
        lock (_lock)
        {
            if (_state != RecorderState.Recording)
                return (false, "Not currently recording.", null, null);

            // Signal the hook thread to exit its message loop and unhook
            _hookThreadRunning = false;
            if (_hookThread != null)
            {
                // Post WM_QUIT to break GetMessage loop
                if (_hookThreadNativeId != 0)
                    PostThreadMessage(_hookThreadNativeId, WM_QUIT, IntPtr.Zero, IntPtr.Zero);
                _hookThread.Join(TimeSpan.FromSeconds(3));
                _hookThread = null;
                _hookThreadNativeId = 0;
            }
            _typingFlushTimer?.Dispose();
            _altSequentialTimer?.Dispose();
            _mouseDelegate = null;
            _keyboardDelegate = null;

            // Flush any pending typing buffer
            FlushTypingBuffer();

            _state = RecorderState.Idle;

            if (_actions.Count == 0)
                return (false, "No actions were recorded.", null, null);

            // Compute wait-before for each action
            ComputeWaits();

            // Build macro
            var displayName = _macroName.Split('/').Last().Replace('-', ' ');
            displayName = char.ToUpper(displayName[0]) + displayName[1..];

            var macro = MacroSerializer.BuildFromRecordedActions(
                displayName,
                $"Recorded macro ({_actions.Count} actions)",
                _actions);

            // Serialize and save
            var yaml = MacroSerializer.ToYaml(macro);
            string? filePath = null;
            try
            {
                filePath = MacroSerializer.SaveToFile(macro, _macroName, _macrosPath);
                Console.Error.WriteLine($"[{DateTime.Now:HH:mm:ss}] Recording saved: {filePath} ({_actions.Count} actions → {macro.Steps.Count} steps)");
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"[{DateTime.Now:HH:mm:ss}] Failed to save recording: {ex.Message}");
                return (true, $"Recorded {_actions.Count} actions but failed to save: {ex.Message}", yaml, null);
            }

            return (true, $"Recorded {_actions.Count} actions → {macro.Steps.Count} macro steps. Saved to {filePath}", yaml, filePath);
        }
    }

    /// <summary>Compute wait durations between actions based on timestamp gaps.</summary>
    private void ComputeWaits()
    {
        for (int i = 1; i < _actions.Count; i++)
        {
            var gap = (_actions[i].Timestamp - _actions[i - 1].Timestamp).TotalSeconds;
            if (gap >= Constants.WaitDetectionThresholdSec)
            {
                _actions[i].WaitBeforeSec = Math.Min(gap, Constants.MaxRecordedWaitSec);
            }
        }
    }

    /// <summary>Check if the foreground window belongs to the target process.</summary>
    private bool IsTargetForeground()
    {
        var targetPid = _engine.ProcessId;
        if (targetPid == null) return false;
        var fgWnd = GetForegroundWindow();
        if (fgWnd == IntPtr.Zero) return false;
        GetWindowThreadProcessId(fgWnd, out uint fgPid);
        return fgPid == (uint)targetPid.Value;
    }

    // =====================================================================
    // Mouse hook
    // =====================================================================

    private IntPtr MouseHookCallback(int nCode, IntPtr wParam, IntPtr lParam)
    {
        if (nCode >= 0 && _state == RecorderState.Recording)
        {
            var msg = (int)wParam;
            if (msg == WM_LBUTTONDOWN || msg == WM_RBUTTONDOWN)
            {
                var hookStruct = Marshal.PtrToStructure<MSLLHOOKSTRUCT>(lParam);
                var x = hookStruct.x;
                var y = hookStruct.y;

                // Filter: only capture clicks on the target process
                if (_engine.IsPointInTargetProcess(x, y))
                {
                    // Flush any pending typing first
                    lock (_lock) { FlushTypingBuffer(); }

                    // Identify element at click point (runs on STA thread)
                    var elementInfo = _engine.ElementFromPoint(x, y);

                    lock (_lock)
                    {
                        _actions.Add(new RecordedAction
                        {
                            Type = msg == WM_LBUTTONDOWN ? RecordedActionType.Click : RecordedActionType.RightClick,
                            Timestamp = DateTime.UtcNow,
                            X = x,
                            Y = y,
                            AutomationId = elementInfo?.AutomationId,
                            ElementName = elementInfo?.Name,
                            ClassName = elementInfo?.ClassName,
                            ControlType = elementInfo?.ControlType,
                        });
                        _lastActionTime = DateTime.UtcNow;
                    }
                }
            }
        }

        return CallNextHookEx(_mouseHook, nCode, wParam, lParam);
    }

    // =====================================================================
    // Keyboard hook
    // =====================================================================

    private IntPtr KeyboardHookCallback(int nCode, IntPtr wParam, IntPtr lParam)
    {
        if (nCode >= 0 && _state == RecorderState.Recording && IsTargetForeground())
        {
            var hookStruct = Marshal.PtrToStructure<KBDLLHOOKSTRUCT>(lParam);
            var vk = (ushort)hookStruct.vkCode;
            var msg = (int)wParam;
            var isKeyDown = msg == WM_KEYDOWN || msg == WM_SYSKEYDOWN;
            var isKeyUp = msg == WM_KEYUP || msg == WM_SYSKEYUP;

            lock (_lock)
            {
                if (IsModifierKey(vk))
                {
                    HandleModifierKey(vk, isKeyDown, isKeyUp);
                }
                else if (isKeyDown)
                {
                    HandleNonModifierKeyDown(vk);
                }
            }
        }

        return CallNextHookEx(_keyboardHook, nCode, wParam, lParam);
    }

    private void HandleModifierKey(ushort vk, bool isKeyDown, bool isKeyUp)
    {
        var normalized = NormalizeModifier(vk);

        if (isKeyDown)
        {
            _modifiersDown.Add(normalized);
            _altReleasedAlone = false;
            _altSequentialTimer?.Dispose();
        }
        else if (isKeyUp)
        {
            _modifiersDown.Remove(normalized);

            // Detect bare Alt release (for ribbon keytips: Alt then F)
            if ((normalized == VK_MENU) && _modifiersDown.Count == 0)
            {
                _altReleasedAlone = true;
                _altSequentialTimer?.Dispose();
                _altSequentialTimer = new Timer(_ =>
                {
                    lock (_lock)
                    {
                        if (_altReleasedAlone)
                        {
                            // Alt was released alone and no follow-up key came — record bare Alt
                            FlushTypingBuffer();
                            _actions.Add(new RecordedAction
                            {
                                Type = RecordedActionType.SendKeys,
                                Timestamp = DateTime.UtcNow,
                                Keys = "Alt"
                            });
                            _altReleasedAlone = false;
                            _lastActionTime = DateTime.UtcNow;
                        }
                    }
                }, null, Constants.AltSequentialTimeoutMs, Timeout.Infinite);
            }
        }
    }

    private void HandleNonModifierKeyDown(ushort vk)
    {
        // Case 1: Alt was released alone, now a follow-up key → sequential "Alt,X"
        if (_altReleasedAlone)
        {
            _altReleasedAlone = false;
            _altSequentialTimer?.Dispose();
            FlushTypingBuffer();

            var keyName = VkToKeyName(vk);
            _actions.Add(new RecordedAction
            {
                Type = RecordedActionType.SendKeys,
                Timestamp = DateTime.UtcNow,
                Keys = $"Alt,{keyName}"
            });
            _lastActionTime = DateTime.UtcNow;
            return;
        }

        // Case 2: Modifiers held → combo like "Ctrl+S"
        if (_modifiersDown.Count > 0)
        {
            FlushTypingBuffer();

            var combo = BuildComboString(vk);
            _actions.Add(new RecordedAction
            {
                Type = RecordedActionType.SendKeys,
                Timestamp = DateTime.UtcNow,
                Keys = combo
            });
            _lastActionTime = DateTime.UtcNow;
            return;
        }

        // Case 3: Bare key press — is it a typeable character or a special key?
        if (IsCharacterKey(vk))
        {
            var ch = VkToChar(vk);
            if (ch.HasValue)
            {
                _typingBuffer.Append(ch.Value);
                _lastTypingTime = DateTime.UtcNow;

                // Reset flush timer
                _typingFlushTimer?.Dispose();
                _typingFlushTimer = new Timer(_ =>
                {
                    lock (_lock) { FlushTypingBuffer(); }
                }, null, Constants.TypingCoalesceMs, Timeout.Infinite);
                return;
            }
        }

        // Case 4: Special key (Enter, Tab, Escape, F-keys, etc.)
        FlushTypingBuffer();
        var specialKeyName = VkToKeyName(vk);
        _actions.Add(new RecordedAction
        {
            Type = RecordedActionType.SendKeys,
            Timestamp = DateTime.UtcNow,
            Keys = specialKeyName
        });
        _lastActionTime = DateTime.UtcNow;
    }

    private void FlushTypingBuffer()
    {
        if (_typingBuffer.Length == 0) return;
        _actions.Add(new RecordedAction
        {
            Type = RecordedActionType.Type,
            Timestamp = _lastTypingTime,
            Text = _typingBuffer.ToString()
        });
        _lastActionTime = _lastTypingTime;
        _typingBuffer.Clear();
        _typingFlushTimer?.Dispose();
    }

    // =====================================================================
    // Key helpers
    // =====================================================================

    private static bool IsModifierKey(ushort vk) =>
        vk is VK_LCONTROL or VK_RCONTROL or VK_LMENU or VK_RMENU
            or VK_LSHIFT or VK_RSHIFT or VK_LWIN or VK_RWIN
            or VK_CONTROL or VK_MENU or VK_SHIFT;

    /// <summary>Normalize left/right modifier variants to a canonical form.</summary>
    private static ushort NormalizeModifier(ushort vk) => vk switch
    {
        VK_LCONTROL or VK_RCONTROL or VK_CONTROL => VK_CONTROL,
        VK_LMENU or VK_RMENU or VK_MENU => VK_MENU,
        VK_LSHIFT or VK_RSHIFT or VK_SHIFT => VK_SHIFT,
        VK_LWIN or VK_RWIN => VK_LWIN,
        _ => vk
    };

    private string BuildComboString(ushort mainVk)
    {
        var parts = new List<string>();
        if (_modifiersDown.Contains(VK_CONTROL)) parts.Add("Ctrl");
        if (_modifiersDown.Contains(VK_MENU)) parts.Add("Alt");
        if (_modifiersDown.Contains(VK_SHIFT)) parts.Add("Shift");
        parts.Add(VkToKeyName(mainVk));
        return string.Join("+", parts);
    }

    /// <summary>Check if a VK code is a typeable character (letters, digits, punctuation).</summary>
    private static bool IsCharacterKey(ushort vk)
    {
        // A-Z
        if (vk >= 0x41 && vk <= 0x5A) return true;
        // 0-9
        if (vk >= 0x30 && vk <= 0x39) return true;
        // Space
        if (vk == 0x20) return true;
        // Numpad 0-9
        if (vk >= 0x60 && vk <= 0x69) return true;
        // OEM keys (semicolon, plus, comma, minus, period, slash, backtick, brackets, backslash, quote)
        if (vk >= 0xBA && vk <= 0xC0) return true;
        if (vk >= 0xDB && vk <= 0xDE) return true;
        // Numpad operators
        if (vk >= 0x6A && vk <= 0x6F) return true;
        return false;
    }

    [DllImport("user32.dll")]
    private static extern int ToUnicode(uint wVirtKey, uint wScanCode, byte[] lpKeyState,
        [Out, MarshalAs(UnmanagedType.LPWStr)] StringBuilder pwszBuff, int cchBuff, uint wFlags);
    [DllImport("user32.dll")]
    private static extern bool GetKeyboardState(byte[] lpKeyState);

    /// <summary>Convert a virtual key to a character using the current keyboard layout.</summary>
    private static char? VkToChar(ushort vk)
    {
        var keyState = new byte[256];
        GetKeyboardState(keyState);
        var sb = new StringBuilder(2);
        int result = ToUnicode(vk, 0, keyState, sb, sb.Capacity, 0);
        if (result == 1) return sb[0];
        return null;
    }

    /// <summary>
    /// Map virtual key code to a human-readable key name for macro YAML.
    /// Uses System.Windows.Input.Key enum via KeyInterop for reliable names.
    /// </summary>
    private static string VkToKeyName(ushort vk)
    {
        // Well-known special keys
        return vk switch
        {
            0x08 => "Back",
            0x09 => "Tab",
            0x0D => "Enter",
            0x1B => "Escape",
            0x20 => "Space",
            0x21 => "PageUp",
            0x22 => "PageDown",
            0x23 => "End",
            0x24 => "Home",
            0x25 => "Left",
            0x26 => "Up",
            0x27 => "Right",
            0x28 => "Down",
            0x2D => "Insert",
            0x2E => "Delete",
            // F1-F24
            >= 0x70 and <= 0x87 => $"F{vk - 0x6F}",
            // Letters A-Z
            >= 0x41 and <= 0x5A => ((char)vk).ToString(),
            // Digits 0-9
            >= 0x30 and <= 0x39 => ((char)vk).ToString(),
            _ => TryKeyInterop(vk) ?? $"VK_{vk:X2}"
        };
    }

    private static string? TryKeyInterop(ushort vk)
    {
        try
        {
            var key = System.Windows.Input.KeyInterop.KeyFromVirtualKey(vk);
            if (key != System.Windows.Input.Key.None)
                return key.ToString();
        }
        catch { }
        return null;
    }

    public void Dispose()
    {
        _hookThreadRunning = false;
        if (_hookThread != null)
        {
            if (_hookThreadNativeId != 0)
                PostThreadMessage(_hookThreadNativeId, WM_QUIT, IntPtr.Zero, IntPtr.Zero);
            _hookThread.Join(TimeSpan.FromSeconds(3));
            _hookThread = null;
        }
        _typingFlushTimer?.Dispose();
        _altSequentialTimer?.Dispose();
    }
}
