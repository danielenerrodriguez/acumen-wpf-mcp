using System.Diagnostics;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;
using Microsoft.Extensions.Logging;
using ModelContextProtocol.Server;
using WpfMcp;

// --server: run elevated, listen on named pipe for UIA commands
if (args.Contains("--server"))
{
    await UiaProxyServer.RunAsync();
    return;
}

// --mcp-connect: non-elevated MCP server that proxies UIA calls to elevated --server
// This is what OpenCode launches.
if (args.Contains("--mcp-connect"))
{
    await RunMcpConnectAsync(args);
    return;
}

// --mcp: standard stdio MCP server (non-elevated, direct mode — for testing only)
if (args.Contains("--mcp"))
{
    var builder = Host.CreateApplicationBuilder(args);

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
                Name = "wpf-uia",
                Version = "1.0.0"
            };
        })
        .WithStdioServerTransport()
        .WithToolsFromAssembly();

    await builder.Build().RunAsync();
    return;
}

// Default (no args / double-click): interactive CLI mode
await CliMode.RunAsync(args);

// =====================================================================
// --mcp-connect implementation
// =====================================================================
async Task RunMcpConnectAsync(string[] cliArgs)
{
    const string MUTEX_NAME = "Global\\WpfMcp_Server_Running";

    // --- Ensure elevated server is running ---
    if (!IsServerRunning(MUTEX_NAME))
    {
        Console.Error.WriteLine("[WPF MCP] Elevated server not running. Launching...");
        Console.Error.WriteLine("[WPF MCP] You may see an elevation prompt. Please approve it.");

        if (!LaunchElevatedServer())
        {
            Console.Error.WriteLine("[WPF MCP] ERROR: Failed to launch elevated server.");
            return;
        }

        Console.Error.WriteLine("[WPF MCP] Waiting for elevated server to start...");
        bool ready = false;
        for (int i = 0; i < 30; i++)
        {
            await Task.Delay(1000);
            if (IsServerRunning(MUTEX_NAME)) { ready = true; break; }
        }

        if (!ready)
        {
            Console.Error.WriteLine("[WPF MCP] ERROR: Server did not start within 30 seconds.");
            return;
        }

        Console.Error.WriteLine("[WPF MCP] Elevated server is running.");
        await Task.Delay(500); // let pipe listener start
    }
    else
    {
        Console.Error.WriteLine("[WPF MCP] Elevated server already running.");
    }

    // --- Connect proxy client to elevated server ---
    var proxy = new UiaProxyClient();
    try
    {
        Console.Error.WriteLine("[WPF MCP] Connecting to elevated server pipe...");
        await proxy.ConnectAsync(timeoutMs: 10000);
        Console.Error.WriteLine("[WPF MCP] Connected!");
    }
    catch (Exception ex)
    {
        Console.Error.WriteLine($"[WPF MCP] ERROR: Could not connect to pipe: {ex.Message}");
        proxy.Dispose();
        return;
    }

    // --- Wire proxy into tools and start MCP stdio server ---
    WpfTools.Proxy = proxy;

    var filteredArgs = cliArgs.Where(a => a != "--mcp-connect").ToArray();
    var builder = Host.CreateApplicationBuilder(filteredArgs);

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
                Name = "wpf-uia",
                Version = "1.0.0"
            };
        })
        .WithStdioServerTransport()
        .WithToolsFromAssembly();

    Console.Error.WriteLine("[WPF MCP] MCP server ready (proxied mode).");

    using var app = builder.Build();
    await app.RunAsync();

    proxy.Dispose();
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

static bool LaunchElevatedServer()
{
    try
    {
        var exePath = Process.GetCurrentProcess().MainModule?.FileName
            ?? System.IO.Path.Combine(AppContext.BaseDirectory, "WpfMcp.exe");

        var psi = new ProcessStartInfo
        {
            FileName = exePath,
            Arguments = "--server",
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
