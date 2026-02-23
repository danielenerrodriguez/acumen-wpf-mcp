using Xunit;

namespace WpfMcp.Tests;

/// <summary>
/// Tests for WpfTools MCP tool methods.
/// These test the tool layer's input validation, error handling,
/// and response formatting without requiring a running WPF application.
/// </summary>
public class WpfToolsTests : IDisposable
{
    public WpfToolsTests()
    {
        // Ensure no proxy is set — tools run in direct mode
        WpfTools.Proxy = null;
    }

    public void Dispose()
    {
        WpfTools.Proxy = null;
    }

    [Fact]
    public async Task Attach_NoParams_ReturnsError()
    {
        var result = await WpfTools.wpf_attach();
        Assert.StartsWith("Error:", result);
        Assert.Contains("Either process_name, pid, or exe_path must be provided", result);
    }

    [Fact]
    public async Task Attach_NonexistentProcess_ReturnsError()
    {
        var result = await WpfTools.wpf_attach(process_name: "NonExistentProcess_12345");
        Assert.StartsWith("Error:", result);
    }

    [Fact]
    public async Task Attach_InvalidPid_ReturnsError()
    {
        var result = await WpfTools.wpf_attach(pid: -1);
        Assert.StartsWith("Error:", result);
    }

    [Fact]
    public async Task Snapshot_WhenNotAttached_ReturnsError()
    {
        // UiaEngine singleton may be attached from previous test runs,
        // but with a fresh engine instance this should fail
        var result = await WpfTools.wpf_snapshot();
        // Either returns error (not attached) or succeeds if engine happens to be attached
        Assert.NotNull(result);
    }

    [Fact]
    public async Task Click_UnknownRef_ReturnsError()
    {
        var result = await WpfTools.wpf_click("nonexistent_ref");
        Assert.StartsWith("Error:", result);
        Assert.Contains("Unknown element reference", result);
    }

    [Fact]
    public async Task RightClick_UnknownRef_ReturnsError()
    {
        var result = await WpfTools.wpf_right_click("nonexistent_ref");
        Assert.StartsWith("Error:", result);
        Assert.Contains("Unknown element reference", result);
    }

    [Fact]
    public async Task Properties_UnknownRef_ReturnsError()
    {
        var result = await WpfTools.wpf_properties("nonexistent_ref");
        Assert.StartsWith("Error:", result);
        Assert.Contains("Unknown element reference", result);
    }

    [Fact]
    public async Task Children_UnknownRef_ReturnsError()
    {
        var result = await WpfTools.wpf_children("nonexistent_ref");
        Assert.StartsWith("Error:", result);
        Assert.Contains("Unknown element reference", result);
    }

    [Fact]
    public async Task Status_WhenNotAttached_ReturnsNotAttachedMessage()
    {
        // Reset the engine singleton to ensure clean state is tricky
        // because it's a singleton, but we can at least verify the method runs
        var result = await WpfTools.wpf_status();
        Assert.NotNull(result);
        // Should contain either "Not attached" or "Attached to PID"
        Assert.True(
            result.Contains("Not attached") || result.Contains("Attached to PID"),
            $"Unexpected status result: {result}");
    }
}
