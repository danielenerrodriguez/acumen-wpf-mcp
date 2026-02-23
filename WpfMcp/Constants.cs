using System.Text.Json;

namespace WpfMcp;

/// <summary>
/// Shared constants used across the WPF MCP server.
/// </summary>
public static class Constants
{
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

    // Shared JSON options
    public static readonly JsonSerializerOptions IndentedJson = new() { WriteIndented = true };
}
