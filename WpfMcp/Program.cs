using System.Diagnostics;
using System.IO;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;
using Microsoft.Extensions.Logging;
using ModelContextProtocol.Server;
using WpfMcp;

// Parse shared arguments (used by both server and client modes)
string? macrosPath = null;
string? shortcutsPath = null;
int? idleTimeoutMinutes = null;
bool noIdle = args.Contains("--no-idle");
for (int i = 0; i < args.Length - 1; i++)
{
    if (args[i] == "--macros-path")
        macrosPath = args[i + 1];
    else if (args[i] == "--shortcuts-path")
        shortcutsPath = args[i + 1];
    else if (args[i] == "--idle-timeout" && int.TryParse(args[i + 1], out var minutes) && minutes >= 0)
        idleTimeoutMinutes = minutes;
}

// Apply idle timeout: --no-idle takes precedence, then --idle-timeout, then default (60)
if (noIdle)
    Constants.ServerIdleTimeoutMinutes = 0;
else if (idleTimeoutMinutes.HasValue)
    Constants.ServerIdleTimeoutMinutes = idleTimeoutMinutes.Value;

// Drag-and-drop: first arg is a .yaml file dropped onto the exe
if (args.Length > 0
    && args[0].EndsWith(".yaml", StringComparison.OrdinalIgnoreCase)
    && File.Exists(args[0]))
{
    await RunDragDropMacroAsync(args[0]);
    return;
}

// --cli: interactive CLI mode for manual testing
if (args.Contains("--cli"))
{
    await CliMode.RunAsync(args, macrosPath);
    return;
}

// --mcp-connect: non-elevated MCP server that proxies UIA calls to elevated --server
// This is what ClaudeCode launches.
if (args.Contains("--mcp-connect"))
{
    await RunMcpConnectAsync(args);
    return;
}

// --mcp: standard stdio MCP server (non-elevated, direct mode — for testing only)
if (args.Contains("--mcp"))
{
    await BuildMcpHost(args, macrosPath).RunAsync();
    return;
}

// "run <path.yaml> [k=v ...]": one-shot macro execution mode
if (args.Length >= 2
    && args[0].Equals("run", StringComparison.OrdinalIgnoreCase)
    && args[1].EndsWith(".yaml", StringComparison.OrdinalIgnoreCase))
{
    await RunDragDropMacroAsync(args[1], args.Skip(2).ToArray());
    return;
}

// --export-all: bulk export all macros as Windows shortcuts, then exit
if (args.Contains("--export-all"))
{
    using var exportEngine = new MacroEngine(macrosPath, enableWatcher: false);
    var results = exportEngine.ExportAllMacros(shortcutsPath, force: args.Contains("--force"));
    int ok = 0, fail = 0;
    foreach (var r in results)
    {
        Console.WriteLine(r.Ok
            ? $"  OK: {r.MacroName} -> {r.ShortcutPath}"
            : $"  FAILED: {r.MacroName} - {r.Message}");
        if (r.Ok) ok++; else fail++;
    }
    Console.WriteLine($"Exported {ok} shortcut(s), {fail} failed.");
    return;
}

// Default (no args / double-click / --server): run elevated server with named pipe + web dashboard
// If not already elevated, re-launch with UAC prompt and exit this instance
if (!IsElevated())
{
    try
    {
        var exePath = Process.GetCurrentProcess().MainModule?.FileName
            ?? Path.Combine(AppContext.BaseDirectory, "WpfMcp.exe");
        var serverArgs = "--server";
        if (macrosPath != null)
            serverArgs += $" --macros-path \"{macrosPath}\"";
        if (noIdle)
            serverArgs += " --no-idle";
        else if (idleTimeoutMinutes.HasValue)
            serverArgs += $" --idle-timeout {idleTimeoutMinutes.Value}";
        Process.Start(new ProcessStartInfo
        {
            FileName = exePath,
            Arguments = serverArgs,
            Verb = "runas",
            UseShellExecute = true
        });
    }
    catch (System.ComponentModel.Win32Exception ex) when (ex.NativeErrorCode == 1223)
    {
        Console.Error.WriteLine("Elevation was denied by the user.");
    }
    return;
}
await UiaProxyServer.RunAsync(macrosPath);

// =====================================================================
// --mcp-connect implementation
// =====================================================================
async Task RunMcpConnectAsync(string[] cliArgs)
{
    var proxy = await EnsureServerAndConnectAsync(Console.Error);
    if (proxy == null) return;

    // --- Wire proxy into tools and start MCP stdio server ---
    WpfTools.Proxy = proxy;

    Console.Error.WriteLine("[WPF MCP] MCP server ready (proxied mode).");

    using var app = BuildMcpHost(cliArgs, macrosPath);
    await app.RunAsync();

    proxy.Dispose();
}

// =====================================================================
// Shared MCP host builder
// =====================================================================
static IHost BuildMcpHost(string[] hostArgs, string? macrosPathOverride = null)
{
    var builder = Host.CreateApplicationBuilder(hostArgs);
    builder.Logging.ClearProviders();
    builder.Logging.AddConsole(options =>
    {
        options.LogToStandardErrorThreshold = LogLevel.Trace;
    });
    builder.Services
        .AddMcpServer(options =>
        {
            options.ServerInfo = new()
            {
                Name = Constants.ServerName,
                Version = Constants.ServerVersion
            };
            options.ServerInstructions = BuildServerInstructions(macrosPathOverride);
        })
        .WithStdioServerTransport()
        .WithToolsFromAssembly()
        .WithResourcesFromAssembly();
    return builder.Build();
}

// =====================================================================
// Build MCP server instructions from macros + knowledge bases
// =====================================================================
static string BuildServerInstructions(string? macrosPathOverride)
{
    try
    {
        using var engine = new MacroEngine(macrosPathOverride, enableWatcher: false);
        var sb = new System.Text.StringBuilder();

        // Macro list
        var macros = engine.List();
        if (macros.Count > 0)
        {
            sb.AppendLine("# Available Macros");
            sb.AppendLine();
            foreach (var m in macros)
            {
                sb.Append($"- **{m.Name}**: {m.Description}");
                if (m.Parameters.Count > 0)
                {
                    var paramList = string.Join(", ", m.Parameters.Select(p =>
                        p.Required ? $"{p.Name} (required)" : $"{p.Name}={p.Default ?? "optional"}"));
                    sb.Append($"  [{paramList}]");
                }
                sb.AppendLine();
            }
            sb.AppendLine();
        }

        // Full knowledge base content
        var knowledgeBases = engine.KnowledgeBases;
        if (knowledgeBases.Count > 0)
        {
            foreach (var kb in knowledgeBases)
            {
                sb.AppendLine($"# Knowledge Base: {kb.ProductName}");
                sb.AppendLine();
                sb.AppendLine(kb.FullContent);
                sb.AppendLine();
            }
        }

        var result = sb.ToString().TrimEnd();
        return result.Length > 0 ? result : "No macros or knowledge bases loaded.";
    }
    catch (Exception ex)
    {
        return $"Failed to load macros/knowledge bases: {ex.Message}";
    }
}

// =====================================================================
// Shared: Ensure elevated server is running and connect via pipe
// =====================================================================
async Task<UiaProxyClient?> EnsureServerAndConnectAsync(TextWriter log)
{
    if (!IsServerRunning(Constants.MutexName))
    {
        log.WriteLine("[WPF MCP] Elevated server not running. Launching...");
        log.WriteLine("[WPF MCP] You may see an elevation prompt. Please approve it.");

        if (!LaunchElevatedServer())
        {
            log.WriteLine("[WPF MCP] ERROR: Failed to launch elevated server.");
            return null;
        }

        log.WriteLine("[WPF MCP] Waiting for elevated server to start...");
        bool ready = false;
        for (int i = 0; i < Constants.ServerStartupTimeoutSeconds; i++)
        {
            await Task.Delay(Constants.ServerStartupPollMs);
            if (IsServerRunning(Constants.MutexName)) { ready = true; break; }
        }

        if (!ready)
        {
            log.WriteLine($"[WPF MCP] ERROR: Server did not start within {Constants.ServerStartupTimeoutSeconds} seconds.");
            return null;
        }

        log.WriteLine("[WPF MCP] Elevated server is running.");
        await Task.Delay(Constants.ServerPostStartDelayMs);
    }
    else
    {
        log.WriteLine("[WPF MCP] Elevated server already running.");
    }

    var proxy = new UiaProxyClient();
    try
    {
        log.WriteLine("[WPF MCP] Connecting to elevated server pipe...");
        await proxy.ConnectAsync(timeoutMs: Constants.PipeConnectTimeoutMs);
        log.WriteLine("[WPF MCP] Connected!");
        return proxy;
    }
    catch (Exception ex)
    {
        log.WriteLine($"[WPF MCP] ERROR: Could not connect to pipe: {ex.Message}");
        proxy.Dispose();
        return null;
    }
}

// =====================================================================
// Helpers
// =====================================================================
static bool IsElevated()
{
    using var identity = System.Security.Principal.WindowsIdentity.GetCurrent();
    var principal = new System.Security.Principal.WindowsPrincipal(identity);
    return principal.IsInRole(System.Security.Principal.WindowsBuiltInRole.Administrator);
}

static bool IsServerRunning(string mutexName)
{
    try
    {
        bool createdNew;
        using var mutex = new Mutex(false, mutexName, out createdNew);
        if (createdNew)
        {
            mutex.ReleaseMutex();
            return false;
        }
        return true;
    }
    catch { return false; }
}

bool LaunchElevatedServer()
{
    try
    {
        var exePath = Process.GetCurrentProcess().MainModule?.FileName
            ?? System.IO.Path.Combine(AppContext.BaseDirectory, "WpfMcp.exe");

        var serverArgs = "--server";
        if (macrosPath != null)
            serverArgs += $" --macros-path \"{macrosPath}\"";
        if (noIdle)
            serverArgs += " --no-idle";
        else if (idleTimeoutMinutes.HasValue)
            serverArgs += $" --idle-timeout {idleTimeoutMinutes.Value}";

        var psi = new ProcessStartInfo
        {
            FileName = exePath,
            Arguments = serverArgs,
            Verb = "runas",
            UseShellExecute = true,
            WindowStyle = ProcessWindowStyle.Normal
        };

        var proc = Process.Start(psi);
        return proc != null;
    }
    catch (System.ComponentModel.Win32Exception ex) when (ex.NativeErrorCode == 1223)
    {
        Console.Error.WriteLine("[WPF MCP] Elevation was denied by the user.");
        return false;
    }
    catch (Exception ex)
    {
        Console.Error.WriteLine($"[WPF MCP] Launch error: {ex.Message}");
        return false;
    }
}

// =====================================================================
// Drag-and-drop YAML macro execution
// =====================================================================
async Task RunDragDropMacroAsync(string yamlPath, string[]? cliParams = null)
{
    Console.WriteLine($"WPF MCP - Running macro: {Path.GetFileName(yamlPath)}");
    Console.WriteLine(new string('=', 50));

    // Read and parse the YAML
    string yamlContent;
    MacroDefinition macro;
    try
    {
        yamlContent = File.ReadAllText(yamlPath);
        macro = YamlHelpers.Deserializer.Deserialize<MacroDefinition>(yamlContent);
    }
    catch (Exception ex)
    {
        Console.WriteLine($"Failed to parse YAML: {ex.Message}");
        WaitForKeypress();
        return;
    }

    if (macro == null || macro.Steps.Count == 0)
    {
        Console.WriteLine("Invalid macro: no steps defined.");
        WaitForKeypress();
        return;
    }

    var displayName = macro.Name ?? Path.GetFileNameWithoutExtension(yamlPath);
    Console.WriteLine($"Macro: {displayName}");
    if (!string.IsNullOrEmpty(macro.Description))
        Console.WriteLine($"Description: {macro.Description}");
    Console.WriteLine($"Steps: {macro.Steps.Count}");
    Console.WriteLine();

    // Parse CLI-provided parameters (key=value format)
    var macroParams = new Dictionary<string, string>();
    if (cliParams != null)
    {
        foreach (var arg in cliParams)
        {
            var eqIdx = arg.IndexOf('=');
            if (eqIdx > 0)
                macroParams[arg[..eqIdx]] = arg[(eqIdx + 1)..];
        }
    }

    // Prompt for any remaining required parameters not provided via CLI
    foreach (var p in macro.Parameters)
    {
        if (p.Required && p.Default == null && !macroParams.ContainsKey(p.Name))
        {
            Console.Write($"  {p.Name}{(string.IsNullOrEmpty(p.Description) ? "" : $" ({p.Description})")}: ");
            var value = Console.ReadLine()?.Trim();
            if (!string.IsNullOrEmpty(value))
                macroParams[p.Name] = value;
        }
    }

    // Ensure elevated server is running and connect
    using var proxy = await EnsureServerAndConnectAsync(Console.Out);
    if (proxy == null)
    {
        WaitForKeypress();
        return;
    }

    // Create a temp log file for real-time progress display
    var logFile = Path.Combine(Path.GetTempPath(), $"wpfmcp-macro-{Guid.NewGuid():N}.log");

    // Execute via proxy
    Console.WriteLine($"Executing '{displayName}'...");
    Console.WriteLine();

    // Start background task to tail the log file for real-time output
    using var logCts = new CancellationTokenSource();
    var logTailState = new LogTailState();
    var logTask = TailLogFileAsync(logFile, logTailState, logCts.Token);

    try
    {
        var proxyArgs = new Dictionary<string, object?>
        {
            ["yaml"] = yamlContent,
            ["parameters"] = macroParams.Count > 0
                ? System.Text.Json.JsonSerializer.Serialize(macroParams)
                : null,
            ["logFile"] = logFile
        };
        var response = await proxy.CallAsync("executeMacroYaml", proxyArgs);

        // Stop tailing and flush remaining lines
        logCts.Cancel();
        try { await logTask; } catch (OperationCanceledException) { }
        FlushLogFile(logFile, logTailState);

        var ok = response.GetProperty("ok").GetBoolean();
        if (ok)
        {
            var result = response.GetProperty("result");
            var steps = result.GetProperty("StepsExecuted").GetInt32();
            var total = result.GetProperty("TotalSteps").GetInt32();
            Console.WriteLine();
            Console.WriteLine($"SUCCESS: {displayName} completed ({steps}/{total} steps)");
            return; // Auto-close on success
        }
        else
        {
            var result = response.GetProperty("result");
            var msg = result.GetProperty("Message").GetString();
            Console.WriteLine();
            Console.WriteLine($"FAILED: {msg}");
            if (result.TryGetProperty("Error", out var err) && err.ValueKind == System.Text.Json.JsonValueKind.String)
                Console.WriteLine($"  Error: {err.GetString()}");
            var steps = result.GetProperty("StepsExecuted").GetInt32();
            var total = result.GetProperty("TotalSteps").GetInt32();
            Console.WriteLine($"  Steps completed: {steps}/{total}");
        }
    }
    catch (Exception ex)
    {
        logCts.Cancel();
        try { await logTask; } catch (OperationCanceledException) { }
        FlushLogFile(logFile, logTailState);
        Console.WriteLine($"ERROR: {ex.Message}");
    }
    finally
    {
        // Clean up temp log file
        try { File.Delete(logFile); } catch { }
    }

    WaitForKeypress(); // Only pause on failure so user can read the error
}

static void WaitForKeypress()
{
    Console.WriteLine();
    Console.WriteLine("Press any key to exit...");
    Console.ReadKey(intercept: true);
}

/// <summary>
/// Background task that tails a log file and prints new lines to the console in real-time.
/// The server writes macro step log messages to this file; we poll and display them as they appear.
/// </summary>
static async Task TailLogFileAsync(string logFile, LogTailState state, CancellationToken ct)
{
    try
    {
        // Wait briefly for the file to be created by the server
        while (!File.Exists(logFile) && !ct.IsCancellationRequested)
            await Task.Delay(100, ct);

        while (!ct.IsCancellationRequested)
        {
            ReadNewLines(logFile, state);
            await Task.Delay(200, ct);
        }
    }
    catch (OperationCanceledException)
    {
        // Expected when the macro completes
    }
}

/// <summary>Reads and prints new lines from the log file starting at the tracked position.</summary>
static void ReadNewLines(string logFile, LogTailState state)
{
    try
    {
        if (!File.Exists(logFile)) return;

        using var fs = new FileStream(logFile, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
        if (fs.Length <= state.LastPosition) return;

        fs.Seek(state.LastPosition, SeekOrigin.Begin);
        using var reader = new StreamReader(fs);
        while (!reader.EndOfStream)
        {
            var line = reader.ReadLine();
            if (line != null)
                Console.WriteLine(line);
        }
        state.LastPosition = fs.Position;
    }
    catch (IOException)
    {
        // File may be briefly locked by the server — retry on next poll
    }
}

/// <summary>
/// Reads and prints any remaining lines from the log file that the tail task may have missed.
/// Called after the macro completes to ensure all log output is displayed.
/// </summary>
static void FlushLogFile(string logFile, LogTailState state)
{
    ReadNewLines(logFile, state);
}

/// <summary>Tracks the file position for the log tail task so the flush can pick up where it left off.</summary>
class LogTailState { public long LastPosition; }
