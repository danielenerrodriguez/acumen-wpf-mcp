using System.Diagnostics;
using System.IO;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;
using Microsoft.Extensions.Logging;
using ModelContextProtocol.Server;
using WpfMcp;

// Parse --macros-path argument (used by both server and client modes)
string? macrosPath = null;
for (int i = 0; i < args.Length - 1; i++)
{
    if (args[i] == "--macros-path")
    {
        macrosPath = args[i + 1];
        break;
    }
}

// Drag-and-drop: first arg is a .yaml file dropped onto the exe
if (args.Length > 0
    && args[0].EndsWith(".yaml", StringComparison.OrdinalIgnoreCase)
    && File.Exists(args[0]))
{
    await RunDragDropMacroAsync(args[0]);
    return;
}

// --server: run elevated, listen on named pipe for UIA commands
if (args.Contains("--server"))
{
    await UiaProxyServer.RunAsync(macrosPath);
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

// Default (no args / double-click): interactive CLI mode
await CliMode.RunAsync(args, macrosPath);

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

    // Execute via proxy
    Console.WriteLine($"Executing '{displayName}'...");
    Console.WriteLine();
    try
    {
        var proxyArgs = new Dictionary<string, object?>
        {
            ["yaml"] = yamlContent,
            ["parameters"] = macroParams.Count > 0
                ? System.Text.Json.JsonSerializer.Serialize(macroParams)
                : null
        };
        var response = await proxy.CallAsync("executeMacroYaml", proxyArgs);

        var ok = response.GetProperty("ok").GetBoolean();
        if (ok)
        {
            var result = response.GetProperty("result");
            var steps = result.GetProperty("StepsExecuted").GetInt32();
            var total = result.GetProperty("TotalSteps").GetInt32();
            Console.WriteLine($"SUCCESS: {displayName} completed ({steps}/{total} steps)");
            return; // Auto-close on success
        }
        else
        {
            var result = response.GetProperty("result");
            var msg = result.GetProperty("Message").GetString();
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
        Console.WriteLine($"ERROR: {ex.Message}");
    }

    WaitForKeypress(); // Only pause on failure so user can read the error
}

static void WaitForKeypress()
{
    Console.WriteLine();
    Console.WriteLine("Press any key to exit...");
    Console.ReadKey(intercept: true);
}
