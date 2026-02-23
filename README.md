# WPF UIA MCP Server

An MCP (Model Context Protocol) server that provides UI Automation access to WPF desktop applications. Built to work with AI coding agents like [OpenCode](https://opencode.ai) and any MCP-compatible client.

Originally built for **Deltek Acumen Fuse**, but designed to work with any WPF application.

## What It Does

This server lets an AI agent interact with a running WPF application through UI Automation:

- **Attach** to a WPF process by name or PID
- **Snapshot** the visual tree (controls, properties, automation IDs)
- **Find** elements by AutomationId, Name, ClassName, or ControlType
- **Find by path** using hierarchical ObjectStore-style paths
- **Click / Right-click** elements
- **Type text** into focused elements
- **Send keyboard shortcuts** (e.g., `Ctrl+S`, `Alt,F` for ribbon keytips)
- **Take screenshots** of the application window
- **Inspect properties** of any cached element

## Why the Proxy Architecture

WPF applications like Fuse require **elevated (admin) privileges** for UI Automation to see the full control tree. Without elevation, UIA only sees a single Win32 TitleBar — none of the WPF content.

MCP clients launch server processes via stdio, which doesn't support elevation (`runas` requires `UseShellExecute = true`, which disables stdio redirection). The solution is a two-process proxy:

```
MCP Client (e.g., OpenCode)
  |  stdio
  v
WpfMcp.exe --mcp-connect     (non-elevated, MCP protocol over stdio)
  |  named pipe (WpfMcp_UIA)
  v
WpfMcp.exe --server           (elevated, executes UIA commands)
  |  System.Windows.Automation
  v
Target WPF Application
```

The `--mcp-connect` process auto-launches the elevated server if it isn't already running. The elevated server persists across client reconnections (5 min idle timeout) and remembers the last attached process for auto-reattach.

## Prerequisites

- Windows 10/11
- .NET 9 SDK (with `net9.0-windows` workload)
- The target WPF application must be running
- Elevation capability (UAC, BeyondTrust, or similar) for the `--server` process

## Building

```
cd C:\WpfMcp
dotnet build -c Release
```

Output: `C:\WpfMcp\WpfMcp\bin\Release\net9.0-windows\WpfMcp.exe`

### Running Tests

```
dotnet test
```

## Usage Modes

### MCP Mode (for AI agents)

Configure your MCP client to launch:

```json
{
  "wpf-uia": {
    "type": "local",
    "command": ["cmd.exe", "/c", "C:\\WpfMcp\\WpfMcp\\bin\\Release\\net9.0-windows\\WpfMcp.exe --mcp-connect"]
  }
}
```

On first use, you'll see a BeyondTrust/UAC prompt to elevate the server process. Approve it once — the elevated server stays running for 5 minutes of idle time.

### CLI Mode (for manual testing)

Double-click the exe or run without arguments:

```
WpfMcp.exe
```

Interactive commands: `attach <name>`, `snapshot`, `find`, `click`, `keys`, `children`, etc.

### Direct MCP Mode (non-elevated, for testing)

```
WpfMcp.exe --mcp
```

Runs a standard MCP stdio server without elevation. Only works if the target app doesn't require elevated UIA access.

## MCP Tools Reference

| Tool | Description |
|------|-------------|
| `wpf_attach` | Attach to a WPF process by name or PID |
| `wpf_status` | Check attachment status (process name, PID, window title) |
| `wpf_snapshot` | Get the UI automation tree (configurable depth) |
| `wpf_find` | Find an element by AutomationId, Name, ClassName, or ControlType |
| `wpf_find_by_path` | Find an element by hierarchical path segments |
| `wpf_children` | List children of an element (or the main window) |
| `wpf_click` | Click an element by reference key |
| `wpf_right_click` | Right-click an element by reference key |
| `wpf_type` | Type text into the focused element |
| `wpf_send_keys` | Send keyboard input (`Ctrl+S`, `Alt,F`, `Enter`, etc.) |
| `wpf_focus` | Bring the application window to the foreground |
| `wpf_screenshot` | Capture the application window as a PNG |
| `wpf_properties` | Get detailed properties of a cached element |

### Element References

`wpf_find`, `wpf_find_by_path`, and `wpf_children` return element references (e.g., `e1`, `e2`). Use these refs with `wpf_click`, `wpf_right_click`, and `wpf_properties`. References are cached on the elevated server and persist across MCP client reconnections.

### Keyboard Input

- **Simultaneous keys**: Use `+` separator — `Ctrl+S`, `Alt+F4`, `Ctrl+Shift+N`
- **Sequential keys**: Use `,` separator — `Alt,F` (press Alt, release, then press F)
- **Single keys**: `Enter`, `Escape`, `Tab`, `F5`, `Delete`

### ObjectStore Paths

`wpf_find_by_path` accepts hierarchical path segments matching the QEAutomation ObjectStore format:

```
SearchProp:ControlType~Custom;SearchProp:AutomationId~uxProjectsView
```

This is useful for navigating deep into the control tree where a simple `wpf_find` might match multiple elements.

## Project Structure

```
C:\WpfMcp\
  WpfMcp.slnx                          Solution file
  nuget.config                          NuGet source configuration
  WpfMcp/
    WpfMcp.csproj                       Main project (.NET 9, WPF)
    Program.cs                          Entry point, mode routing, auto-launch logic
    Tools.cs                            MCP tool definitions (proxied and direct modes)
    UiaEngine.cs                        Core UI Automation engine (STA thread, SendInput)
    UiaProxy.cs                         Proxy client/server for named pipe communication
    CliMode.cs                          Interactive CLI for manual testing
  WpfMcp.Tests/
    WpfMcp.Tests.csproj                 xUnit test project
```

## Troubleshooting

**"Access to the path is denied" on pipe connection**
The elevated server creates the pipe with an ACL that allows authenticated users. If you still get this error, ensure the `--server` process is actually elevated.

**Snapshot only shows TitleBar / FrameworkId="Win32"**
The server process is not elevated. Kill it and re-launch — approve the elevation prompt.

**"Not attached to any process"**
Call `wpf_attach` with the process name. After the first attach, the server remembers the process and auto-reattaches on reconnection.

**Server shuts down after 5 minutes**
The elevated server has an idle timeout. It auto-relaunches when the next `--mcp-connect` client starts.

**Build errors about Infragistics NuGet source**
The `nuget.config` file clears external sources. Make sure it's present in the project root.
