using System.IO;
using Microsoft.AspNetCore.Builder;
using Microsoft.AspNetCore.Hosting;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Logging;
using WpfMcp.Web;
using WpfMcp.Web.Components;

namespace WpfMcp;

/// <summary>
/// Starts a Kestrel web server hosting the Blazor Server dashboard.
/// Runs alongside the named pipe server in the elevated --server process.
/// </summary>
internal static class WebServer
{
    private static WebApplication? _app;

    /// <summary>
    /// Start the Blazor Server dashboard on a background thread.
    /// Returns once the server is listening. Does not block.
    /// </summary>
    public static async Task StartAsync(IAppState appState, int port, CancellationToken ct = default)
    {
        var ready = new TaskCompletionSource();

        _ = Task.Run(async () =>
        {
            try
            {
                // Use the exe directory as content root so ASP.NET Core can find assemblies
                var exeDir = Path.GetDirectoryName(Environment.ProcessPath)
                    ?? AppContext.BaseDirectory;

                var builder = WebApplication.CreateBuilder(new WebApplicationOptions
                {
                    ApplicationName = typeof(App).Assembly.GetName().Name,
                    ContentRootPath = exeDir,
                    WebRootPath = Path.Combine(exeDir, "wwwroot")
                });

                // Configure Kestrel — bind all interfaces so WSL browsers can connect
                builder.WebHost.UseUrls($"http://*:{port}");

                // Keep console logging so errors are visible in the server terminal,
                // but set minimum level to Warning to avoid noise
                builder.Logging.ClearProviders();
                builder.Logging.AddConsole();
                builder.Logging.SetMinimumLevel(Microsoft.Extensions.Logging.LogLevel.Warning);

                // Register Blazor Server services
                builder.Services.AddRazorComponents()
                    .AddInteractiveServerComponents();

                // Register IAppState as a singleton so Blazor components can inject it
                builder.Services.AddSingleton(appState);

                var app = builder.Build();

                app.UseStaticFiles();
                app.UseAntiforgery();

                // Map Blazor components — App lives in WpfMcp.Web RCL
                // No AddAdditionalAssemblies needed since App IS in the RCL assembly
                app.MapRazorComponents<App>()
                    .AddInteractiveServerRenderMode();

                _app = app;

                // Signal that startup config is complete
                app.Lifetime.ApplicationStarted.Register(() => ready.TrySetResult());

                // Stop the web server when the cancellation token fires
                ct.Register(() => app.StopAsync().ConfigureAwait(false));

                await app.RunAsync();
            }
            catch (Exception ex)
            {
                ready.TrySetException(ex);
            }
        }, ct);

        // Wait for the server to actually start listening
        await ready.Task;
    }

    /// <summary>Gracefully stop the web server.</summary>
    public static async Task StopAsync()
    {
        if (_app is not null)
        {
            await _app.StopAsync();
            _app = null;
        }
    }
}
