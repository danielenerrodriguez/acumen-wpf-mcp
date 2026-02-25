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

    /// <summary>Default constructor — call <see cref="ConnectAsync"/> to establish the pipe.</summary>
    public UiaProxyClient() { }

    /// <summary>
    /// Internal constructor for testing — accepts a pre-connected pipe stream.
    /// Avoids the need for reflection-based injection in tests.
    /// </summary>
    internal UiaProxyClient(NamedPipeClientStream pipe)
    {
        _pipe = pipe;
        _reader = new StreamReader(pipe);
        _writer = new StreamWriter(pipe) { AutoFlush = true };
    }

    public async Task ConnectAsync(int timeoutMs = 5000)
    {
        _pipe = new NamedPipeClientStream(".", Constants.PipeName, PipeDirection.InOut, PipeOptions.Asynchronous);
        await _pipe.ConnectAsync(timeoutMs);
        _reader = new StreamReader(_pipe);
        _writer = new StreamWriter(_pipe) { AutoFlush = true };
    }

    public async Task<JsonElement> CallAsync(
        string method,
        Dictionary<string, object?>? args = null,
        CancellationToken cancellation = default)
    {
        if (_pipe == null || !_pipe.IsConnected)
            throw new InvalidOperationException("Not connected to elevated UIA server");

        await _lock.WaitAsync(cancellation);
        try
        {
            var request = new Dictionary<string, object?>
            {
                ["method"] = method,
                ["args"] = args ?? new Dictionary<string, object?>()
            };

            var json = JsonSerializer.Serialize(request);
            await _writer!.WriteLineAsync(json.AsMemory(), cancellation);

            var response = await _reader!.ReadLineAsync(cancellation);
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
        try { _lock.Dispose(); } catch { }
    }
}

/// <summary>
/// Elevated side: listens on a named pipe, executes UIA commands via UiaEngine,
/// returns results as JSON.
/// </summary>
public static class UiaProxyServer
{
    private static readonly string _lastClientFile =
        Path.Combine(
            Path.GetDirectoryName(Environment.ProcessPath) ?? AppContext.BaseDirectory,
            "last_client.txt");

    private static readonly ElementCache _cache = new();
    private static readonly Lazy<MacroEngine> _macroEngine = new(() => new MacroEngine());
    private static readonly SemaphoreSlim _commandLock = new(1, 1);
    private static string? _macrosPath;
    private static int _clientCount;

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

    private static string? GetStringArg(JsonElement args, string name) =>
        args.TryGetProperty(name, out var v) && v.ValueKind == JsonValueKind.String ? v.GetString() : null;

    private static int? GetIntArg(JsonElement args, string name) =>
        args.TryGetProperty(name, out var v) && v.ValueKind == JsonValueKind.Number ? v.GetInt32() : (int?)null;

    public static async Task RunAsync(string? macrosPath = null)
    {
        _macrosPath = macrosPath;
        Console.WriteLine("========================================");
        Console.WriteLine("  WPF UIA Server (Elevated)");
        Console.WriteLine("========================================");
        Console.WriteLine();
        Console.WriteLine("Provides UI Automation access to WPF apps.");
        Console.WriteLine("Elevation is required to read the visual");
        Console.WriteLine("tree of protected applications.");
        Console.WriteLine();
        Console.WriteLine($"Pipe: {Constants.PipeName}");
        Console.WriteLine($"Auto-shutdown: {Constants.ServerIdleTimeoutMinutes} min idle");
        Console.WriteLine("Status: READY");
        Console.WriteLine();

        // Acquire mutex to signal we're running
        using var mutex = new Mutex(true, Constants.MutexName);

        // Shared security descriptor — all pipe instances must use identical ACL.
        var pipeSecurity = new PipeSecurity();
        pipeSecurity.AddAccessRule(new PipeAccessRule(
            new SecurityIdentifier(WellKnownSidType.AuthenticatedUserSid, null),
            PipeAccessRights.ReadWrite,
            AccessControlType.Allow));
        pipeSecurity.AddAccessRule(new PipeAccessRule(
            new SecurityIdentifier(WellKnownSidType.BuiltinAdministratorsSid, null),
            PipeAccessRights.FullControl,
            AccessControlType.Allow));

        // Track active client handlers so we only idle-shutdown when none are connected
        var activeClients = new List<Task>();

        while (true)
        {
            try
            {
                // Multiple instances (up to 5) allow concurrent clients (e.g., MCP + drag-and-drop).
                var pipe = NamedPipeServerStreamAcl.Create(
                    Constants.PipeName, PipeDirection.InOut, 5,
                    PipeTransmissionMode.Byte, PipeOptions.Asynchronous,
                    inBufferSize: 0, outBufferSize: 0, pipeSecurity);

                // Wait for client with idle timeout (only when no clients are connected)
                using var idleCts = new CancellationTokenSource(TimeSpan.FromMinutes(Constants.ServerIdleTimeoutMinutes));
                try
                {
                    await pipe.WaitForConnectionAsync(idleCts.Token);
                }
                catch (OperationCanceledException)
                {
                    pipe.Dispose();
                    // Clean up completed client tasks
                    activeClients.RemoveAll(t => t.IsCompleted);
                    if (activeClients.Count > 0)
                    {
                        // Still have active clients — keep waiting
                        continue;
                    }
                    Console.WriteLine($"[{DateTime.Now:HH:mm:ss}] Idle timeout. Shutting down.");
                    return;
                }

                var clientId = Interlocked.Increment(ref _clientCount);
                Console.WriteLine($"[{DateTime.Now:HH:mm:ss}] Client #{clientId} connected.");

                // Auto-reattach to last known process on first connection
                if (clientId == 1) TryAutoReattach();

                // Handle this client in a background task, loop back to accept more
                var clientTask = Task.Run(() => HandleClientAsync(pipe, clientId));
                activeClients.Add(clientTask);

                // Clean up completed tasks
                activeClients.RemoveAll(t => t.IsCompleted);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"[{DateTime.Now:HH:mm:ss}] Error: {ex.Message}");
                await Task.Delay(Constants.ServerErrorRetryMs);
            }
        }
    }

    private static async Task HandleClientAsync(NamedPipeServerStream pipe, int clientId)
    {
        try
        {
            using (pipe)
            using (var reader = new StreamReader(pipe))
            using (var writer = new StreamWriter(pipe) { AutoFlush = true })
            {
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

                        // Serialize command execution to avoid concurrent UIA calls
                        await _commandLock.WaitAsync();
                        try
                        {
                            response = await ExecuteCommand(method, args);
                        }
                        finally
                        {
                            _commandLock.Release();
                        }
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
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"[{DateTime.Now:HH:mm:ss}] Client #{clientId} error: {ex.Message}");
        }

        Console.WriteLine($"[{DateTime.Now:HH:mm:ss}] Client #{clientId} disconnected.");
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
                    case Constants.Commands.Attach:
                    {
                        var name = GetStringArg(args, "processName");
                        var pid = GetIntArg(args, "pid");
                        var result = pid.HasValue ? engine.AttachByPid(pid.Value) : engine.Attach(name!);
                        if (result.success && name != null)
                            SaveLastClient(name);
                        return Json(result.success, result.message);
                    }
                    case Constants.Commands.Snapshot:
                    {
                        var depth = GetIntArg(args, "maxDepth") ?? 3;
                        var result = engine.GetSnapshot(depth);
                        return Json(result.success, result.tree);
                    }
                    case Constants.Commands.Children:
                    {
                        System.Windows.Automation.AutomationElement? parent = null;
                        var key = GetStringArg(args, "refKey");
                        if (key != null)
                        {
                            if (!_cache.TryGet(key, out parent))
                                return Json(false, $"Unknown ref '{key}'");
                        }
                        var children = engine.GetChildElements(parent);
                        var items = new List<Dictionary<string, string>>();
                        foreach (var child in children)
                        {
                            var refKey = _cache.Add(child);
                            items.Add(new Dictionary<string, string>
                            {
                                ["ref"] = refKey,
                                ["desc"] = engine.FormatElement(child)
                            });
                        }
                        return JsonSerializer.Serialize(new { ok = true, result = items });
                    }
                    case Constants.Commands.Find:
                    {
                        var automationId = GetStringArg(args, "automationId");
                        var name = GetStringArg(args, "name");
                        var className = GetStringArg(args, "className");
                        var controlType = GetStringArg(args, "controlType");
                        var result = engine.FindElement(automationId, name, className, controlType);
                        if (result.success && result.element != null)
                        {
                            var refKey = _cache.Add(result.element);
                            var props = engine.GetElementProperties(result.element);
                            return JsonSerializer.Serialize(new { ok = true, refKey, desc = result.message, properties = props });
                        }
                        return Json(false, result.message);
                    }
                    case Constants.Commands.FindByPath:
                    {
                        var pathArr = args.GetProperty("path");
                        var segments = new List<string>();
                        foreach (var seg in pathArr.EnumerateArray())
                            segments.Add(seg.GetString()!);
                        var result = engine.FindElementByPath(segments);
                        if (result.success && result.element != null)
                        {
                            var refKey = _cache.Add(result.element);
                            var props = engine.GetElementProperties(result.element);
                            return JsonSerializer.Serialize(new { ok = true, refKey, desc = result.message, properties = props });
                        }
                        return Json(false, result.message);
                    }
                    case Constants.Commands.Click:
                    {
                        var key = args.GetProperty("refKey").GetString()!;
                        if (!_cache.TryGet(key, out var el)) return Json(false, $"Unknown ref '{key}'");
                        var result = engine.ClickElement(el!);
                        return Json(result.success, result.message);
                    }
                    case Constants.Commands.RightClick:
                    {
                        var key = args.GetProperty("refKey").GetString()!;
                        if (!_cache.TryGet(key, out var el)) return Json(false, $"Unknown ref '{key}'");
                        var result = engine.RightClickElement(el!);
                        return Json(result.success, result.message);
                    }
                    case Constants.Commands.Type:
                    {
                        var text = args.GetProperty("text").GetString()!;
                        var result = engine.TypeText(text);
                        return Json(result.success, result.message);
                    }
                    case Constants.Commands.SendKeys:
                    {
                        var keys = args.GetProperty("keys").GetString()!;
                        var result = engine.SendKeyboardShortcut(keys);
                        return Json(result.success, result.message);
                    }
                    case Constants.Commands.SetValue:
                    {
                        var key = args.GetProperty("refKey").GetString()!;
                        if (!_cache.TryGet(key, out var el)) return Json(false, $"Unknown ref '{key}'");
                        var value = args.GetProperty("value").GetString()!;
                        var result = engine.SetElementValue(el!, value);
                        return Json(result.success, result.message);
                    }
                    case Constants.Commands.GetValue:
                    {
                        var key = args.GetProperty("refKey").GetString()!;
                        if (!_cache.TryGet(key, out var el)) return Json(false, $"Unknown ref '{key}'");
                        var result = engine.GetElementValue(el!);
                        if (result.success)
                            return JsonSerializer.Serialize(new { ok = true, result = result.value });
                        return Json(false, result.message);
                    }
                    case Constants.Commands.Focus:
                    {
                        var result = engine.FocusWindow();
                        return Json(result.success, result.message);
                    }
                    case Constants.Commands.FileDialog:
                    {
                        var filePath = args.GetProperty("filePath").GetString()!;
                        var result = engine.FileDialogSetPath(filePath);
                        return Json(result.success, result.message);
                    }
                    case Constants.Commands.Screenshot:
                    {
                        var result = engine.TakeScreenshot();
                        if (result.success)
                            return JsonSerializer.Serialize(new { ok = true, result = result.message, base64 = result.base64 });
                        return Json(false, result.message);
                    }
                    case Constants.Commands.Properties:
                    {
                        var key = args.GetProperty("refKey").GetString()!;
                        if (!_cache.TryGet(key, out var el)) return Json(false, $"Unknown ref '{key}'");
                        var props = engine.GetElementProperties(el!);
                        return JsonSerializer.Serialize(new { ok = true, result = props });
                    }
                    case Constants.Commands.Status:
                    {
                        return JsonSerializer.Serialize(new
                        {
                            ok = true,
                            attached = engine.IsAttached,
                            windowTitle = engine.WindowTitle,
                            pid = engine.ProcessId
                        });
                    }
                    case Constants.Commands.MacroList:
                    {
                        var macroEng = _macroEngine.Value;
                        var macros = macroEng.List();
                        var loadErrors = macroEng.LoadErrors;
                        var knowledgeBases = macroEng.KnowledgeBases.Select(kb => new
                        {
                            productName = kb.ProductName,
                            filePath = kb.FilePath,
                            summary = kb.Summary
                        }).ToList();
                        return JsonSerializer.Serialize(new { ok = true, result = macros, loadErrors, knowledgeBases });
                    }
                    case Constants.Commands.Macro:
                    {
                        var macroName = GetStringArg(args, "name");
                        if (string.IsNullOrEmpty(macroName))
                            return Json(false, "Macro name is required");

                        var parsedParams = new Dictionary<string, string>();
                        var paramsJson = GetStringArg(args, "parameters");
                        if (!string.IsNullOrEmpty(paramsJson))
                        {
                            try
                            {
                                var p = JsonSerializer.Deserialize<Dictionary<string, string>>(paramsJson);
                                if (p != null) parsedParams = p;
                            }
                            catch (Exception ex)
                            {
                                return Json(false, $"Invalid parameters JSON: {ex.Message}");
                            }
                        }

                        var macroResult = _macroEngine.Value.ExecuteAsync(
                            macroName, parsedParams, engine, _cache).GetAwaiter().GetResult();
                        return JsonSerializer.Serialize(new { ok = macroResult.Success, result = macroResult });
                    }
                    case Constants.Commands.ExecuteMacroYaml:
                    {
                        var yamlContent = GetStringArg(args, "yaml");
                        if (string.IsNullOrEmpty(yamlContent))
                            return Json(false, "YAML content is required");

                        MacroDefinition macroDef;
                        try
                        {
                            macroDef = YamlHelpers.Deserializer.Deserialize<MacroDefinition>(yamlContent);
                        }
                        catch (Exception ex)
                        {
                            return Json(false, $"Failed to parse YAML: {ex.Message}");
                        }

                        if (macroDef == null || macroDef.Steps.Count == 0)
                            return Json(false, "Invalid macro: no steps defined");

                        var execParams = new Dictionary<string, string>();
                        var execParamsJson = GetStringArg(args, "parameters");
                        if (!string.IsNullOrEmpty(execParamsJson))
                        {
                            try
                            {
                                var p = JsonSerializer.Deserialize<Dictionary<string, string>>(execParamsJson);
                                if (p != null) execParams = p;
                            }
                            catch (Exception ex)
                            {
                                return Json(false, $"Invalid parameters JSON: {ex.Message}");
                            }
                        }

                        var displayName = macroDef.Name ?? "inline-macro";
                        var execResult = _macroEngine.Value.ExecuteDefinitionAsync(
                            macroDef, displayName, execParams, engine, _cache).GetAwaiter().GetResult();
                        return JsonSerializer.Serialize(new { ok = execResult.Success, result = execResult });
                    }
                    case Constants.Commands.Launch:
                    {
                        var exePath = GetStringArg(args, "exePath");
                        if (string.IsNullOrEmpty(exePath))
                            return Json(false, "exe_path is required");
                        var arguments = GetStringArg(args, "arguments");
                        var workingDir = GetStringArg(args, "workingDirectory");
                        var ifNotRunning = !(args.TryGetProperty("ifNotRunning", out var inr)
                            && inr.ValueKind == JsonValueKind.False);
                        var timeout = GetIntArg(args, "timeout") ?? Constants.DefaultLaunchTimeoutSec;

                        var result = engine.LaunchAndAttachAsync(
                            exePath, arguments, workingDir, ifNotRunning, timeout).GetAwaiter().GetResult();
                        if (result.success)
                        {
                            var exeName = Path.GetFileNameWithoutExtension(exePath);
                            SaveLastClient(exeName);
                        }
                        return Json(result.success, result.message);
                    }
                    case Constants.Commands.WaitForWindow:
                    {
                        var titleContains = GetStringArg(args, "titleContains");
                        var automationId = GetStringArg(args, "automationId");
                        var name = GetStringArg(args, "name");
                        var controlType = GetStringArg(args, "controlType");
                        var timeout = GetIntArg(args, "timeout") ?? Constants.DefaultLaunchTimeoutSec;
                        var pollMs = GetIntArg(args, "pollMs") ?? Constants.DefaultWindowReadyPollMs;

                        var result = engine.WaitForWindowReadyAsync(
                            titleContains, automationId, name, controlType,
                            timeout, pollMs).GetAwaiter().GetResult();
                        return Json(result.success, result.message);
                    }
                    case Constants.Commands.SaveMacro:
                    {
                        var saveName = GetStringArg(args, "name");
                        if (string.IsNullOrEmpty(saveName))
                            return Json(false, "Macro name is required");

                        var saveDesc = GetStringArg(args, "description") ?? "";
                        var stepsJson = GetStringArg(args, "steps");
                        if (string.IsNullOrEmpty(stepsJson))
                            return Json(false, "Steps JSON is required");

                        var paramsJson = GetStringArg(args, "parameters");
                        var saveTimeout = GetIntArg(args, "timeout") ?? 30;
                        var saveForce = args.TryGetProperty("force", out var sf) &&
                            sf.ValueKind == JsonValueKind.True;

                        // Parse steps
                        List<Dictionary<string, object>> parsedSteps;
                        try
                        {
                            parsedSteps = JsonHelpers.ParseJsonArray(stepsJson);
                        }
                        catch (Exception ex)
                        {
                            return Json(false, $"Invalid steps JSON: {ex.Message}");
                        }

                        // Parse parameters
                        List<Dictionary<string, object>>? parsedParams = null;
                        if (!string.IsNullOrEmpty(paramsJson))
                        {
                            try
                            {
                                parsedParams = JsonHelpers.ParseJsonArray(paramsJson);
                            }
                            catch (Exception ex)
                            {
                                return Json(false, $"Invalid parameters JSON: {ex.Message}");
                            }
                        }

                        // Get attached process name
                        if (!engine.IsAttached)
                            return Json(false, "Not attached to any process. Call wpf_attach first.");
                        var processName = engine.ProcessName;

                        var saveResult = _macroEngine.Value.SaveMacro(
                            saveName, saveDesc, parsedSteps, parsedParams,
                            saveTimeout, saveForce, processName);

                        if (saveResult.Ok)
                            return JsonSerializer.Serialize(new
                            {
                                ok = true,
                                filePath = saveResult.FilePath,
                                macroName = saveResult.MacroName,
                                message = saveResult.Message
                            });
                        return Json(false, saveResult.Message);
                    }
                    case Constants.Commands.ExportMacro:
                    {
                        var exportName = GetStringArg(args, "name");
                        if (string.IsNullOrEmpty(exportName))
                            return Json(false, "Macro name is required");

                        var outputPath = GetStringArg(args, "outputPath");
                        var exportForce = args.TryGetProperty("force", out var ef) &&
                            ef.ValueKind == JsonValueKind.True;

                        var exportResult = _macroEngine.Value.ExportMacro(
                            exportName, outputPath, exportForce);

                        if (exportResult.Ok)
                            return JsonSerializer.Serialize(new
                            {
                                ok = true,
                                shortcutPath = exportResult.ShortcutPath,
                                macroName = exportResult.MacroName,
                                message = exportResult.Message
                            });
                        return Json(false, exportResult.Message);
                    }
                    case Constants.Commands.ExportAllMacros:
                    {
                        var outputPath = GetStringArg(args, "outputPath");
                        var exportForce = args.TryGetProperty("force", out var eaf) &&
                            eaf.ValueKind == JsonValueKind.True;

                        var results = _macroEngine.Value.ExportAllMacros(outputPath, exportForce);
                        var resultList = results.Select(r => new
                        {
                            ok = r.Ok,
                            shortcutPath = r.ShortcutPath,
                            macroName = r.MacroName,
                            message = r.Message
                        }).ToList();

                        return JsonSerializer.Serialize(new { ok = true, results = resultList });
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
