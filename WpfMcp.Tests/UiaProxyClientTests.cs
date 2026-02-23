using System.IO.Pipes;
using System.Text.Json;
using Xunit;

namespace WpfMcp.Tests;

public class UiaProxyClientTests
{
    [Fact]
    public void IsConnected_DefaultsFalse()
    {
        using var client = new UiaProxyClient();
        Assert.False(client.IsConnected);
    }

    [Fact]
    public async Task CallAsync_WhenNotConnected_Throws()
    {
        using var client = new UiaProxyClient();
        var ex = await Assert.ThrowsAsync<InvalidOperationException>(
            () => client.CallAsync("status"));
        Assert.Contains("Not connected", ex.Message);
    }

    [Fact]
    public async Task Dispose_WhenNotConnected_DoesNotThrow()
    {
        var client = new UiaProxyClient();
        client.Dispose();
        // Should not throw
        await Task.CompletedTask;
    }
}
