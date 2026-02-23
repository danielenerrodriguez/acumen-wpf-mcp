using System.IO;
using System.Text.Json;

namespace WpfMcp;

/// <summary>
/// Shared constants used across the WPF MCP server.
/// </summary>
public static class Constants
{
    /// <summary>
    /// Resolve the macros folder path. Priority: explicit path > env var > macros/ next to exe.
    /// Uses Environment.ProcessPath (actual exe location on disk) rather than
    /// AppContext.BaseDirectory (which points to a temp extraction dir for single-file publish).
    /// </summary>
    public static string ResolveMacrosPath(string? explicitPath = null) =>
        explicitPath
        ?? Environment.GetEnvironmentVariable("WPFMCP_MACROS_PATH")
        ?? Path.Combine(
            Path.GetDirectoryName(Environment.ProcessPath) ?? AppContext.BaseDirectory,
            "macros");

    // Named pipe and mutex
    public const string PipeName = "WpfMcp_UIA";
    public const string MutexName = "Global\\WpfMcp_Server_Running";
    public const string ServerName = "wpf-uia";
    public const string ServerVersion = "1.0.0";

    // Timeouts (milliseconds unless noted)
    public const int ServerStartupTimeoutSeconds = 30;
    public const int ServerStartupPollMs = 1000;
    public const int ServerPostStartDelayMs = 500;
    public const int PipeConnectTimeoutMs = 10000;
    public const int ServerIdleTimeoutMinutes = 5;
    public const int ServerErrorRetryMs = 1000;

    // UIA interaction delays
    public const int FocusDelayMs = 300;
    public const int PreKeyDelayMs = 200;
    public const int SequentialKeyDelayMs = 150;
    public const int PreTypeDelayMs = 100;
    public const int PerCharDelayMs = 20;
    public const int PreClickDelayMs = 50;
    public const int PostClickDelayMs = 100;
    public const int DefaultSnapshotDepth = 3;
    public const int MaxWalkDepth = 10;

    // Macro defaults
    public const int DefaultMacroTimeoutSec = 60;
    public const int DefaultStepTimeoutSec = 5;
    public const double DefaultRetryIntervalSec = 1.0;
    public const int MacroReloadDebounceMs = 500;

    // Launch / window readiness
    public const int DefaultLaunchTimeoutSec = 60;
    public const int DefaultWindowReadyPollMs = 2000;
    public const int ProcessMainWindowPollMs = 500;

    // Recording
    public const int TypingCoalesceMs = 300;
    public const double WaitDetectionThresholdSec = 1.5;
    public const double MaxRecordedWaitSec = 10.0;
    public const int AltSequentialTimeoutMs = 500;
    public const int DefaultRecordingStepTimeoutSec = 5;

    // Shared JSON options
    public static readonly JsonSerializerOptions IndentedJson = new() { WriteIndented = true };
}
