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

    /// <summary>
    /// Reproduces the REAL production scenario: MCP tools do NOT pass CancellationToken.
    /// When no CT is passed, CallAsync blocks on ReadLineAsync with CancellationToken.None.
    /// The MCP framework abandons the Task but it keeps running. Meanwhile, the framework
    /// may fire the next tool call on a new Task. Since no CT means no cancellation,
    /// the lock is held until the server finally responds. The second call blocks on the lock.
    /// 
    /// If the MCP transport itself dies (broken pipe), all bets are off.
    /// </summary>
    [Fact]
    public async Task CallAsync_NoCancellationToken_SlowServer_BlocksSubsequentCalls()
    {
        var (server, reader, writer, client) = await SetupPipeAsync(
            nameof(CallAsync_NoCancellationToken_SlowServer_BlocksSubsequentCalls));
        try
        {
            // Call #1: no CancellationToken (matches real Tools.cs usage)
            var call1Task = client.CallAsync("slowMacro", new() { ["name"] = "import-xer" });

            // Server reads request #1
            var request1Json = await reader.ReadLineAsync();
            Assert.Contains("slowMacro", request1Json);

            // Server is slow — doesn't respond yet.
            // MCP framework would abandon the tool handler here (timeout).
            // But call1Task is still blocked on ReadLineAsync with no way to cancel.

            // Call #2: another tool call fired by MCP framework on a different Task
            using var call2Cts = new CancellationTokenSource(TimeSpan.FromSeconds(2));
            var call2Task = client.CallAsync("status", cancellation: call2Cts.Token);

            // call2 blocks on _lock.WaitAsync because call1 holds it.
            // It should time out after 2 seconds.
            await Assert.ThrowsAnyAsync<OperationCanceledException>(() => call2Task);

            // CONFIRMED: Without CancellationToken, slow server calls block ALL subsequent calls.
            // This is the exact production deadlock scenario.
        }
        finally
        {
            // Unblock call1 so test can exit
            try { await writer.WriteLineAsync("{\"ok\":true}"); } catch { }
            // Wait for call1 to drain
            try { await Task.WhenAny(client.CallAsync("noop"), Task.Delay(1000)); } catch { }
            client.Dispose();
            try { reader.Dispose(); } catch { }
            try { writer.Dispose(); } catch { }
            try { server.Dispose(); } catch { }
        }
    }

    /// <summary>
    /// Proves that when call #1 eventually completes (server responds late),
    /// the pipe stays in sync because call1 consumes its own response.
    /// The real problem is NOT desync — it's the deadlock while waiting.
    /// </summary>
    [Fact]
    public async Task CallAsync_SlowServer_LateResponse_PipeStaysInSync()
    {
        var (server, reader, writer, client) = await SetupPipeAsync(
            nameof(CallAsync_SlowServer_LateResponse_PipeStaysInSync));
        try
        {
            // Call #1: client sends request, server is slow
            var call1Task = client.CallAsync("slowMacro", new() { ["name"] = "import-xer" });

            // Server reads request #1
            var request1Json = await reader.ReadLineAsync();
            Assert.Contains("slowMacro", request1Json);

            // Server responds late (after "timeout" but call1 is still waiting)
            await Task.Delay(200); // simulate delay
            await writer.WriteLineAsync("{\"ok\":true,\"result\":\"Macro completed\"}");

            // Call #1 gets its response and releases the lock
            var response1 = await call1Task;
            Assert.True(response1.GetProperty("ok").GetBoolean());

            // Call #2: should work normally since call1 consumed its own response
            var call2Task = client.CallAsync("status");
            var request2Json = await reader.ReadLineAsync();
            Assert.Contains("status", request2Json);
            await writer.WriteLineAsync("{\"ok\":true,\"attached\":true,\"process\":\"Fuse\"}");

            var response2 = await call2Task;
            Assert.True(response2.GetProperty("attached").GetBoolean());
            Assert.Equal("Fuse", response2.GetProperty("process").GetString());
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
