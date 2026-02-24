using System.IO;
using System.IO.Pipes;
using System.Text.Json;
using Xunit;

namespace WpfMcp.Tests;

/// <summary>
/// Tests the named pipe JSON protocol between UiaProxyClient and a mock server.
/// Each test creates its own uniquely-named pipe to avoid collisions.
/// </summary>
public class UiaProxyProtocolTests
{
    private static async Task<(NamedPipeServerStream server, StreamReader reader, StreamWriter writer, UiaProxyClient client)>
        SetupPipeAsync(string testName)
    {
        var pipeName = $"WpfMcp_Test_{testName}_{Guid.NewGuid():N}";

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

    [Fact]
    public async Task SendsCorrectRequestFormat_WithArgs()
    {
        var (server, reader, writer, client) = await SetupPipeAsync(nameof(SendsCorrectRequestFormat_WithArgs));
        try
        {
            var callTask = client.CallAsync("attach", new() { ["processName"] = "Fuse", ["pid"] = null });

            var requestJson = await reader.ReadLineAsync();
            Assert.NotNull(requestJson);

            var request = JsonSerializer.Deserialize<JsonElement>(requestJson);
            Assert.Equal("attach", request.GetProperty("method").GetString());

            var args = request.GetProperty("args");
            Assert.Equal("Fuse", args.GetProperty("processName").GetString());
            Assert.Equal(JsonValueKind.Null, args.GetProperty("pid").ValueKind);

            await writer.WriteLineAsync("{\"ok\":true,\"result\":\"Attached\"}");
            var response = await callTask;
            Assert.True(response.GetProperty("ok").GetBoolean());
        }
        finally
        {
            client.Dispose();
            try { reader.Dispose(); } catch { }
            try { writer.Dispose(); } catch { }
            try { server.Dispose(); } catch { }
        }
    }

    [Fact]
    public async Task SendsEmptyArgs_WhenNoArgsPassed()
    {
        var (server, reader, writer, client) = await SetupPipeAsync(nameof(SendsEmptyArgs_WhenNoArgsPassed));
        try
        {
            var callTask = client.CallAsync("status");

            var requestJson = await reader.ReadLineAsync();
            var request = JsonSerializer.Deserialize<JsonElement>(requestJson!);
            Assert.Equal("status", request.GetProperty("method").GetString());

            var args = request.GetProperty("args");
            Assert.Equal(JsonValueKind.Object, args.ValueKind);
            Assert.Empty(args.EnumerateObject().ToList());

            await writer.WriteLineAsync("{\"ok\":true,\"attached\":false}");
            await callTask;
        }
        finally
        {
            client.Dispose();
            try { reader.Dispose(); } catch { }
            try { writer.Dispose(); } catch { }
            try { server.Dispose(); } catch { }
        }
    }

    [Fact]
    public async Task ParsesSuccessResponse()
    {
        var (server, reader, writer, client) = await SetupPipeAsync(nameof(ParsesSuccessResponse));
        try
        {
            var callTask = client.CallAsync("focus");
            await reader.ReadLineAsync();
            await writer.WriteLineAsync("{\"ok\":true,\"result\":\"Window focused\"}");

            var response = await callTask;
            Assert.True(response.GetProperty("ok").GetBoolean());
            Assert.Equal("Window focused", response.GetProperty("result").GetString());
        }
        finally
        {
            client.Dispose();
            try { reader.Dispose(); } catch { }
            try { writer.Dispose(); } catch { }
            try { server.Dispose(); } catch { }
        }
    }

    [Fact]
    public async Task ParsesErrorResponse()
    {
        var (server, reader, writer, client) = await SetupPipeAsync(nameof(ParsesErrorResponse));
        try
        {
            var callTask = client.CallAsync("snapshot");
            await reader.ReadLineAsync();
            await writer.WriteLineAsync("{\"ok\":false,\"error\":\"Not attached\"}");

            var response = await callTask;
            Assert.False(response.GetProperty("ok").GetBoolean());
            Assert.Equal("Not attached", response.GetProperty("error").GetString());
        }
        finally
        {
            client.Dispose();
            try { reader.Dispose(); } catch { }
            try { writer.Dispose(); } catch { }
            try { server.Dispose(); } catch { }
        }
    }

    [Fact]
    public async Task ParsesComplexFindResponse()
    {
        var (server, reader, writer, client) = await SetupPipeAsync(nameof(ParsesComplexFindResponse));
        try
        {
            var callTask = client.CallAsync("find", new() { ["automationId"] = "uxRibbon" });
            await reader.ReadLineAsync();
            await writer.WriteLineAsync(
                "{\"ok\":true,\"refKey\":\"e1\",\"desc\":\"[ToolBar] uxRibbon\",\"properties\":{\"ControlType\":\"ToolBar\"}}");

            var response = await callTask;
            Assert.True(response.GetProperty("ok").GetBoolean());
            Assert.Equal("e1", response.GetProperty("refKey").GetString());
            Assert.Equal("ToolBar", response.GetProperty("properties").GetProperty("ControlType").GetString());
        }
        finally
        {
            client.Dispose();
            try { reader.Dispose(); } catch { }
            try { writer.Dispose(); } catch { }
            try { server.Dispose(); } catch { }
        }
    }

    [Fact]
    public async Task ThrowsIOException_WhenServerDisconnects()
    {
        var (server, reader, writer, client) = await SetupPipeAsync(nameof(ThrowsIOException_WhenServerDisconnects));
        try
        {
            var callTask = client.CallAsync("status");
            await reader.ReadLineAsync();
            server.Disconnect();

            await Assert.ThrowsAsync<IOException>(() => callTask);
        }
        finally
        {
            client.Dispose();
            try { reader.Dispose(); } catch { }
            try { writer.Dispose(); } catch { }
            try { server.Dispose(); } catch { }
        }
    }

    [Fact]
    public async Task IsConnected_AfterConnect_ReturnsTrue()
    {
        var (server, reader, writer, client) = await SetupPipeAsync(nameof(IsConnected_AfterConnect_ReturnsTrue));
        try
        {
            Assert.True(client.IsConnected);
        }
        finally
        {
            client.Dispose();
            try { reader.Dispose(); } catch { }
            try { writer.Dispose(); } catch { }
            try { server.Dispose(); } catch { }
        }
    }
}
