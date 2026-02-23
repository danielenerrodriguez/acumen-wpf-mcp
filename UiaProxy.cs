using System.IO;
using System.IO.Pipes;
using System.Security.AccessControl;
using System.Security.Principal;
using System.Text.Json;

namespace WpfMcp;

/// <summary>
/// Simple JSON-RPC-like protocol over named pipes between the non-elevated
/// MCP server and the elevated UIA helper.
/// 
/// Request:  {"method":"attach","args":{"processName":"Fuse"}}
/// Response: {"ok":true,"result":"Attached to..."}  or  {"ok":false,"error":"..."}
/// </summary>
public class UiaProxyClient : IDisposable
{
    private NamedPipeClientStream? _pipe;
    private StreamReader? _reader;
    private StreamWriter? _writer;
    private readonly SemaphoreSlim _lock = new(1, 1);

    public bool IsConnected => _pipe?.IsConnected ?? false;

    public async Task ConnectAsync(int timeoutMs = 5000)
    {
        _pipe = new NamedPipeClientStream(".", "WpfMcp_UIA", PipeDirection.InOut, PipeOptions.Asynchronous);
        await _pipe.ConnectAsync(timeoutMs);
        _reader = new StreamReader(_pipe);
        _writer = new StreamWriter(_pipe) { AutoFlush = true };
    }

    public async Task<JsonElement> CallAsync(string method, Dictionary<string, object?>? args = null)
    {
        if (_pipe == null || !_pipe.IsConnected)
            throw new InvalidOperationException("Not connected to elevated UIA server");

        await _lock.WaitAsync();
        try
        {
            var request = new Dictionary<string, object?>
            {
                ["method"] = method,
                ["args"] = args ?? new Dictionary<string, object?>()
            };

            var json = JsonSerializer.Serialize(request);
            await _writer!.WriteLineAsync(json);

            var response = await _reader!.ReadLineAsync();
            if (response == null)
                throw new IOException("Server closed connection");

            return JsonSerializer.Deserialize<JsonElement>(response);
        }
        finally
        {
            _lock.Release();
        }
    }

    public void Dispose()
    {
        try { _reader?.Dispose(); } catch { }
        try { _writer?.Dispose(); } catch { }
        try { _pipe?.Dispose(); } catch { }
    }
}

/// <summary>
/// Elevated side: listens on a named pipe, executes UIA commands via UiaEngine,
/// returns results as JSON.
/// </summary>
public static class UiaProxyServer
{
    private static readonly string _lastClientFile =
        System.IO.Path.Combine(AppContext.BaseDirectory, "last_client.txt");

    /// <summary>Save last attached process name so we can auto-reattach.</summary>
    private static void SaveLastClient(string processName)
    {
        try { File.WriteAllText(_lastClientFile, processName); }
        catch { /* best effort */ }
    }

    /// <summary>Try to auto-reattach to the last known process.</summary>
    private static void TryAutoReattach()
    {
        var engine = UiaEngine.Instance;
        if (engine.IsAttached) return;

        try
        {
            if (!File.Exists(_lastClientFile)) return;
            var name = File.ReadAllText(_lastClientFile).Trim();
            if (string.IsNullOrEmpty(name)) return;

            Console.WriteLine($"[{DateTime.Now:HH:mm:ss}] Auto-reattaching to '{name}'...");
            var result = engine.Attach(name);
            if (result.success)
                Console.WriteLine($"[{DateTime.Now:HH:mm:ss}] Auto-reattached: {result.message}");
            else
                Console.WriteLine($"[{DateTime.Now:HH:mm:ss}] Auto-reattach failed: {result.message}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"[{DateTime.Now:HH:mm:ss}] Auto-reattach error: {ex.Message}");
        }
    }

    public static async Task RunAsync()
    {
        Console.WriteLine("========================================");
        Console.WriteLine("  WPF UIA Server (Elevated)");
        Console.WriteLine("========================================");
        Console.WriteLine();
        Console.WriteLine("Provides UI Automation access to WPF apps.");
        Console.WriteLine("Elevation is required to read the visual");
        Console.WriteLine("tree of protected applications.");
        Console.WriteLine();
        Console.WriteLine("Pipe: WpfMcp_UIA");
        Console.WriteLine("Auto-shutdown: 5 min idle");
        Console.WriteLine("Status: READY");
        Console.WriteLine();

        // Acquire mutex to signal we're running
        using var mutex = new Mutex(true, "Global\\WpfMcp_Server_Running");

        while (true)
        {
            try
            {
                // Create pipe with ACL that allows non-elevated processes to connect.
                // Without this, an elevated server's pipe is only accessible to admins.
                var pipeSecurity = new PipeSecurity();
                pipeSecurity.AddAccessRule(new PipeAccessRule(
                    new SecurityIdentifier(WellKnownSidType.AuthenticatedUserSid, null),
                    PipeAccessRights.ReadWrite,
                    AccessControlType.Allow));

                using var pipe = NamedPipeServerStreamAcl.Create(
                    "WpfMcp_UIA", PipeDirection.InOut, 1,
                    PipeTransmissionMode.Byte, PipeOptions.Asynchronous,
                    inBufferSize: 0, outBufferSize: 0, pipeSecurity);

                // Wait for client with idle timeout
                using var idleCts = new CancellationTokenSource(TimeSpan.FromMinutes(5));
                try
                {
                    await pipe.WaitForConnectionAsync(idleCts.Token);
                }
                catch (OperationCanceledException)
                {
                    Console.WriteLine($"[{DateTime.Now:HH:mm:ss}] Idle timeout. Shutting down.");
                    return;
                }

                Console.WriteLine($"[{DateTime.Now:HH:mm:ss}] Client connected.");

                // Auto-reattach to last known process on new connection
                TryAutoReattach();

                using var reader = new StreamReader(pipe);
                using var writer = new StreamWriter(pipe) { AutoFlush = true };

                // Process commands
                while (pipe.IsConnected)
                {
                    string? line;
                    try
                    {
                        line = await reader.ReadLineAsync();
                    }
                    catch { break; }

                    if (line == null) break;

                    string response;
                    try
                    {
                        var request = JsonSerializer.Deserialize<JsonElement>(line);
                        var method = request.GetProperty("method").GetString()!;
                        var args = request.GetProperty("args");
                        response = await ExecuteCommand(method, args);
                    }
                    catch (Exception ex)
                    {
                        response = JsonSerializer.Serialize(new { ok = false, error = ex.Message });
                    }

                    try
                    {
                        await writer.WriteLineAsync(response);
                    }
                    catch { break; }
                }

                Console.WriteLine($"[{DateTime.Now:HH:mm:ss}] Client disconnected. Waiting...");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"[{DateTime.Now:HH:mm:ss}] Error: {ex.Message}");
                await Task.Delay(1000);
            }
        }
    }

    private static readonly Dictionary<string, System.Windows.Automation.AutomationElement> _elementCache = new();
    private static int _refCounter = 0;

    private static string CacheElement(System.Windows.Automation.AutomationElement el)
    {
        var key = $"e{++_refCounter}";
        _elementCache[key] = el;
        return key;
    }

    static Task<string> ExecuteCommand(string method, JsonElement args)
    {
        var engine = UiaEngine.Instance;

        return Task.Run(() =>
        {
            try
            {
                switch (method)
                {
                    case "attach":
                    {
                        var name = args.TryGetProperty("processName", out var pn) && pn.ValueKind != JsonValueKind.Null ? pn.GetString() : null;
                        var pid = args.TryGetProperty("pid", out var pp) && pp.ValueKind != JsonValueKind.Null ? pp.GetInt32() : (int?)null;
                        var result = pid.HasValue ? engine.AttachByPid(pid.Value) : engine.Attach(name!);
                        if (result.success && name != null)
                            SaveLastClient(name);
                        return Json(result.success, result.message);
                    }
                    case "snapshot":
                    {
                        var depth = args.TryGetProperty("maxDepth", out var d) && d.ValueKind == JsonValueKind.Number ? d.GetInt32() : 3;
                        var result = engine.GetSnapshot(depth);
                        return Json(result.success, result.tree);
                    }
                    case "children":
                    {
                        System.Windows.Automation.AutomationElement? parent = null;
                        if (args.TryGetProperty("refKey", out var rk) && rk.ValueKind == JsonValueKind.String && rk.GetString() is string key)
                        {
                            if (!_elementCache.TryGetValue(key, out parent))
                                return Json(false, $"Unknown ref '{key}'");
                        }
                        var children = engine.GetChildElements(parent);
                        var items = new List<Dictionary<string, string>>();
                        foreach (var child in children)
                        {
                            var refKey = CacheElement(child);
                            items.Add(new Dictionary<string, string>
                            {
                                ["ref"] = refKey,
                                ["desc"] = engine.FormatElement(child)
                            });
                        }
                        return JsonSerializer.Serialize(new { ok = true, result = items });
                    }
                    case "find":
                    {
                        var automationId = args.TryGetProperty("automationId", out var ai) && ai.ValueKind == JsonValueKind.String ? ai.GetString() : null;
                        var name = args.TryGetProperty("name", out var n) && n.ValueKind == JsonValueKind.String ? n.GetString() : null;
                        var className = args.TryGetProperty("className", out var cn) && cn.ValueKind == JsonValueKind.String ? cn.GetString() : null;
                        var controlType = args.TryGetProperty("controlType", out var ct) && ct.ValueKind == JsonValueKind.String ? ct.GetString() : null;
                        var result = engine.FindElement(automationId, name, className, controlType);
                        if (result.success && result.element != null)
                        {
                            var refKey = CacheElement(result.element);
                            var props = engine.GetElementProperties(result.element);
                            return JsonSerializer.Serialize(new { ok = true, refKey, desc = result.message, properties = props });
                        }
                        return Json(false, result.message);
                    }
                    case "findByPath":
                    {
                        var pathArr = args.GetProperty("path");
                        var segments = new List<string>();
                        foreach (var seg in pathArr.EnumerateArray())
                            segments.Add(seg.GetString()!);
                        var result = engine.FindElementByPath(segments);
                        if (result.success && result.element != null)
                        {
                            var refKey = CacheElement(result.element);
                            var props = engine.GetElementProperties(result.element);
                            return JsonSerializer.Serialize(new { ok = true, refKey, desc = result.message, properties = props });
                        }
                        return Json(false, result.message);
                    }
                    case "click":
                    {
                        var key = args.GetProperty("refKey").GetString()!;
                        if (!_elementCache.TryGetValue(key, out var el)) return Json(false, $"Unknown ref '{key}'");
                        var result = engine.ClickElement(el);
                        return Json(result.success, result.message);
                    }
                    case "rightClick":
                    {
                        var key = args.GetProperty("refKey").GetString()!;
                        if (!_elementCache.TryGetValue(key, out var el)) return Json(false, $"Unknown ref '{key}'");
                        var result = engine.RightClickElement(el);
                        return Json(result.success, result.message);
                    }
                    case "type":
                    {
                        var text = args.GetProperty("text").GetString()!;
                        var result = engine.TypeText(text);
                        return Json(result.success, result.message);
                    }
                    case "sendKeys":
                    {
                        var keys = args.GetProperty("keys").GetString()!;
                        var result = engine.SendKeyboardShortcut(keys);
                        return Json(result.success, result.message);
                    }
                    case "focus":
                    {
                        var result = engine.FocusWindow();
                        return Json(result.success, result.message);
                    }
                    case "screenshot":
                    {
                        var result = engine.TakeScreenshot();
                        if (result.success)
                            return JsonSerializer.Serialize(new { ok = true, result = result.message, base64 = result.base64 });
                        return Json(false, result.message);
                    }
                    case "properties":
                    {
                        var key = args.GetProperty("refKey").GetString()!;
                        if (!_elementCache.TryGetValue(key, out var el)) return Json(false, $"Unknown ref '{key}'");
                        var props = engine.GetElementProperties(el);
                        return JsonSerializer.Serialize(new { ok = true, result = props });
                    }
                    case "status":
                    {
                        return JsonSerializer.Serialize(new
                        {
                            ok = true,
                            attached = engine.IsAttached,
                            windowTitle = engine.WindowTitle,
                            pid = engine.ProcessId
                        });
                    }
                    default:
                        return Json(false, $"Unknown method: {method}");
                }
            }
            catch (Exception ex)
            {
                return Json(false, ex.Message);
            }
        });
    }

    static string Json(bool ok, string msg) =>
        JsonSerializer.Serialize(ok ? new { ok, result = msg } : (object)new { ok, error = msg });
}
