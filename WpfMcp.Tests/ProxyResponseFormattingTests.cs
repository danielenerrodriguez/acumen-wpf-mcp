using System.IO;
using System.IO.Pipes;
using System.Text.Json;
using Xunit;

namespace WpfMcp.Tests;

/// <summary>
/// Tests WpfTools response formatting when operating in proxied mode.
/// Uses a mock pipe server to send controlled responses and verifies
/// the tool methods format them correctly for MCP clients.
/// </summary>
public class ProxyResponseFormattingTests
{
    private static async Task<(NamedPipeServerStream server, StreamReader reader, StreamWriter writer, UiaProxyClient client)>
        SetupPipeAsync(string testName)
    {
        var pipeName = $"WpfMcp_Fmt_{testName}_{Guid.NewGuid():N}";

        var server = new NamedPipeServerStream(pipeName,
            PipeDirection.InOut, 1, PipeTransmissionMode.Byte, PipeOptions.Asynchronous);

        var clientPipe = new NamedPipeClientStream(".", pipeName, PipeDirection.InOut, PipeOptions.Asynchronous);

        var serverConnect = server.WaitForConnectionAsync();
        await clientPipe.ConnectAsync(5000);
        await serverConnect;

        var client = new UiaProxyClient(clientPipe);

        var reader = new StreamReader(server);
        var writer = new StreamWriter(server) { AutoFlush = true };

        return (server, reader, writer, client);
    }

    private static void Cleanup(NamedPipeServerStream server, StreamReader reader, StreamWriter writer, UiaProxyClient client)
    {
        WpfTools.Proxy = null;
        client.Dispose();
        try { reader.Dispose(); } catch { }
        try { writer.Dispose(); } catch { }
        try { server.Dispose(); } catch { }
    }

    [Fact]
    public async Task Attach_FormatsSuccessResponse()
    {
        var (server, reader, writer, client) = await SetupPipeAsync(nameof(Attach_FormatsSuccessResponse));
        WpfTools.Proxy = client;
        try
        {
            var task = WpfTools.wpf_attach(process_name: "Fuse");
            await reader.ReadLineAsync();
            await writer.WriteLineAsync("{\"ok\":true,\"result\":\"Attached to Fuse (PID 1234)\"}");

            var result = await task;
            Assert.StartsWith("OK:", result);
            Assert.Contains("Attached to Fuse", result);
        }
        finally { Cleanup(server, reader, writer, client); }
    }

    [Fact]
    public async Task Attach_FormatsErrorResponse()
    {
        var (server, reader, writer, client) = await SetupPipeAsync(nameof(Attach_FormatsErrorResponse));
        WpfTools.Proxy = client;
        try
        {
            var task = WpfTools.wpf_attach(process_name: "Missing");
            await reader.ReadLineAsync();
            await writer.WriteLineAsync("{\"ok\":false,\"error\":\"No process found\"}");

            var result = await task;
            Assert.StartsWith("Error:", result);
            Assert.Contains("No process found", result);
        }
        finally { Cleanup(server, reader, writer, client); }
    }

    [Fact]
    public async Task Snapshot_FormatsTreeOutput()
    {
        var (server, reader, writer, client) = await SetupPipeAsync(nameof(Snapshot_FormatsTreeOutput));
        WpfTools.Proxy = client;
        try
        {
            var task = WpfTools.wpf_snapshot(max_depth: 2);
            await reader.ReadLineAsync();
            await writer.WriteLineAsync("{\"ok\":true,\"result\":\"[Window] Name=\\\"Test\\\"\"}");

            var result = await task;
            Assert.StartsWith("OK:", result);
            Assert.Contains("[Window]", result);
        }
        finally { Cleanup(server, reader, writer, client); }
    }

    [Fact]
    public async Task Find_FormatsRefKeyAndProperties()
    {
        var (server, reader, writer, client) = await SetupPipeAsync(nameof(Find_FormatsRefKeyAndProperties));
        WpfTools.Proxy = client;
        try
        {
            var task = WpfTools.wpf_find(automation_id: "uxRibbon");
            await reader.ReadLineAsync();
            await writer.WriteLineAsync(
                "{\"ok\":true,\"refKey\":\"e5\",\"desc\":\"[ToolBar] uxRibbon\",\"properties\":{\"ControlType\":\"ToolBar\",\"Name\":\"Ribbon\"}}");

            var result = await task;
            Assert.Contains("[e5]", result);
            Assert.Contains("uxRibbon", result);
            Assert.Contains("ToolBar", result);
        }
        finally { Cleanup(server, reader, writer, client); }
    }

    [Fact]
    public async Task Children_FormatsChildList()
    {
        var (server, reader, writer, client) = await SetupPipeAsync(nameof(Children_FormatsChildList));
        WpfTools.Proxy = client;
        try
        {
            var task = WpfTools.wpf_children();
            await reader.ReadLineAsync();
            await writer.WriteLineAsync(
                "{\"ok\":true,\"result\":[{\"ref\":\"e1\",\"desc\":\"[Button] Save\"},{\"ref\":\"e2\",\"desc\":\"[Menu] File\"}]}");

            var result = await task;
            Assert.Contains("Found 2 children", result);
            Assert.Contains("[e1]", result);
            Assert.Contains("[e2]", result);
            Assert.Contains("[Button] Save", result);
        }
        finally { Cleanup(server, reader, writer, client); }
    }

    [Fact]
    public async Task Status_FormatsAttachedState()
    {
        var (server, reader, writer, client) = await SetupPipeAsync(nameof(Status_FormatsAttachedState));
        WpfTools.Proxy = client;
        try
        {
            var task = WpfTools.wpf_status();
            await reader.ReadLineAsync();
            await writer.WriteLineAsync(
                "{\"ok\":true,\"attached\":true,\"windowTitle\":\"Workbook1 - Deltek Acumen\",\"pid\":35992}");

            var result = await task;
            Assert.Contains("Attached to PID 35992", result);
            Assert.Contains("Workbook1 - Deltek Acumen", result);
        }
        finally { Cleanup(server, reader, writer, client); }
    }

    [Fact]
    public async Task Status_FormatsDetachedState()
    {
        var (server, reader, writer, client) = await SetupPipeAsync(nameof(Status_FormatsDetachedState));
        WpfTools.Proxy = client;
        try
        {
            var task = WpfTools.wpf_status();
            await reader.ReadLineAsync();
            await writer.WriteLineAsync("{\"ok\":true,\"attached\":false,\"windowTitle\":null,\"pid\":null}");

            var result = await task;
            Assert.Contains("Not attached", result);
        }
        finally { Cleanup(server, reader, writer, client); }
    }

    [Fact]
    public async Task Screenshot_FormatsBase64Length()
    {
        var (server, reader, writer, client) = await SetupPipeAsync(nameof(Screenshot_FormatsBase64Length));
        WpfTools.Proxy = client;
        try
        {
            var task = WpfTools.wpf_screenshot();
            await reader.ReadLineAsync();
            await writer.WriteLineAsync(
                "{\"ok\":true,\"result\":\"Screenshot (800x600)\",\"base64\":\"iVBORw0KGgo=\"}");

            var result = await task;
            Assert.StartsWith("OK:", result);
            Assert.Contains("Screenshot", result);
            Assert.Contains("base64 PNG", result);
        }
        finally { Cleanup(server, reader, writer, client); }
    }

    [Fact]
    public async Task SendKeys_FormatsSuccess()
    {
        var (server, reader, writer, client) = await SetupPipeAsync(nameof(SendKeys_FormatsSuccess));
        WpfTools.Proxy = client;
        try
        {
            var task = WpfTools.wpf_send_keys("Ctrl+S");
            await reader.ReadLineAsync();
            await writer.WriteLineAsync("{\"ok\":true,\"result\":\"Sent keys: Ctrl+S\"}");

            var result = await task;
            Assert.StartsWith("OK:", result);
            Assert.Contains("Ctrl+S", result);
        }
        finally { Cleanup(server, reader, writer, client); }
    }
}
