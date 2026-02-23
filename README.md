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

## Architecture

```
┌─────────────────────────────────────────────────────────────────┐
│  MCP Client (OpenCode, Claude Desktop, etc.)                    │
│  Launches server via stdio                                      │
└──────────────────────────┬──────────────────────────────────────┘
                           │ stdio (JSON-RPC)
                           v
┌─────────────────────────────────────────────────────────────────┐
│  WpfMcp.exe --mcp-connect              (non-elevated process)   │
│                                                                 │
│  Program.cs         Mode routing, auto-launches elevated server │
│  Tools.cs           15 MCP tool definitions (incl. macros)      │
│  UiaProxyClient     Sends JSON requests over named pipe         │
│  ElementCache       Thread-safe ref cache (e1, e2, ...)         │
│  Constants          Shared config (pipe name, timeouts, etc.)   │
└──────────────────────────┬──────────────────────────────────────┘
                           │ named pipe (WpfMcp_UIA)
                           │ JSON protocol: {"method":"...","args":{...}}
                           v
┌─────────────────────────────────────────────────────────────────┐
│  WpfMcp.exe --server                   (elevated process)       │
│                                                                 │
│  UiaProxyServer     Listens on pipe, dispatches to UiaEngine    │
│  UiaEngine          Singleton, dedicated STA thread for UIA     │
│    ├─ Attach        Process.GetProcessesByName → FromHandle     │
│    ├─ Snapshot      TreeWalker.RawViewWalker recursive walk     │
│    ├─ Find          PropertyCondition + FindFirst/descendants   │
│    ├─ Click         SetCursorPos + mouse_event (via P/Invoke)   │
│    ├─ SendKeys      SendInput API (modifiers + key combos)      │
│    ├─ TypeText      VkKeyScan per char + keybd_event            │
│    └─ Screenshot    GetWindowRect + Graphics.CopyFromScreen     │
│  ElementCache       Server-side ref cache (persists across      │
│                     client reconnections)                       │
└──────────────────────────┬──────────────────────────────────────┘
                           │ System.Windows.Automation (COM/UIA)
                           v
┌─────────────────────────────────────────────────────────────────┐
│  Target WPF Application (e.g., Deltek Acumen Fuse)              │
│  Automation peers exposed via WPF's built-in UIA support        │
└─────────────────────────────────────────────────────────────────┘
```

### Key Design Decisions

- **STA thread** — All UIA calls run on a dedicated STA thread (`UIA-STA`). WPF automation peers require STA COM marshaling; without it, only the Win32 window layer is visible.
- **Lazy singleton** — `UiaEngine` uses `Lazy<T>` with a private constructor. Only one instance exists per process.
- **Reflection-based lookups** — `ControlType` matching uses a dictionary built via reflection on `ControlType`'s static fields, avoiding a manual mapping that would need updating when new types are added.
- **WPF Key enum** — Virtual key parsing uses `System.Windows.Input.Key` + `KeyInterop.VirtualKeyFromKey()` instead of a hardcoded VK table, supporting all 200+ keys automatically.
- **Pipe ACL** — The elevated server creates the named pipe with `PipeSecurity` granting `ReadWrite` to `AuthenticatedUserSid`, allowing the non-elevated client to connect.
- **Auto-reattach** — The server saves the last process name to `last_client.txt` and reattaches on new client connections, so the AI agent doesn't need to re-attach after reconnecting.

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

### CI/CD

A GitHub Actions workflow (`.github/workflows/build-release.yml`) runs on every push to `master`:

1. Builds and runs all tests on `windows-latest`
2. Publishes a self-contained single-file exe (no .NET runtime required)
3. Creates a GitHub Release with `WpfMcp.exe` attached

Pull requests to `master` trigger build and test only (no release).

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
| `wpf_macro` | Execute a named macro with optional parameters |
| `wpf_macro_list` | List all available macros and their parameters |

### Element References

`wpf_find`, `wpf_find_by_path`, and `wpf_children` return element references (e.g., `e1`, `e2`). Use these refs with `wpf_click`, `wpf_right_click`, and `wpf_properties`. References are cached on the elevated server and persist across MCP client reconnections.

### Keyboard Input

- **Simultaneous keys**: Use `+` separator — `Ctrl+S`, `Alt+F4`, `Ctrl+Shift+N`
- **Sequential keys**: Use `,` separator — `Alt,F` (press Alt, release, then press F)
- **Single keys**: `Enter`, `Escape`, `Tab`, `F5`, `Delete`
- **All keys**: Supports all `System.Windows.Input.Key` enum names plus common aliases (`Esc`, `Del`, `Backspace`, `PgUp`, `PgDn`, `Ins`)

### ObjectStore Paths

`wpf_find_by_path` accepts hierarchical path segments matching the QEAutomation ObjectStore format:

```
SearchProp:ControlType~Custom;SearchProp:AutomationId~uxProjectsView
```

This is useful for navigating deep into the control tree where a simple `wpf_find` might match multiple elements.

## Macros

Macros are reusable, YAML-defined step sequences that automate common UI workflows. They support parameters, timeouts, retry logic for find operations, and nested macro calls.

### Folder Structure

```
macros/
  acumen-fuse/
    open-file-menu.yaml
    find-projects-view.yaml
    import-file.yaml
  another-app/
    some-workflow.yaml
```

Macros are discovered from the `macros/` folder next to the exe (copied at build time), or from a custom path via the `WPFMCP_MACROS_PATH` environment variable. Macro names are derived from the relative path without extension, using forward slashes (e.g., `acumen-fuse/import-file`).

### YAML Schema

```yaml
name: Human-readable display name
description: What this macro does
timeout: 30                         # Max seconds for entire macro (default: 60)

parameters:                         # Optional
  - name: filePath
    description: Full path to the file
    required: true
  - name: format
    description: Output format
    required: false
    default: "pdf"

steps:
  - action: focus                   # Bring app window to foreground
  - action: send_keys
    keys: "Alt,F"                   # Keyboard input
  - action: wait
    seconds: 0.5                    # Pause between steps
  - action: find
    name: Import                    # Find element by properties
    control_type: MenuItem
    save_as: importMenuItem         # Save ref for later steps
    timeout: 5                      # Per-step timeout (retries until found)
    retry_interval: 2               # Seconds between retries
  - action: click
    ref: importMenuItem             # Click a saved element
  - action: type
    text: "{{filePath}}"            # Parameter substitution
  - action: macro
    macro_name: acumen-fuse/open-file-menu  # Nested macro call
    params:
      someParam: "{{filePath}}"     # Pass params to nested macro
```

### Supported Actions

| Action | Description | Key Properties |
|--------|-------------|----------------|
| `focus` | Bring application window to foreground | — |
| `attach` | Attach to a process | `process_name`, `pid` |
| `find` | Find element with retry loop | `automation_id`, `name`, `class_name`, `control_type`, `save_as`, `timeout`, `retry_interval` |
| `find_by_path` | Find element by hierarchical path | `path` (list), `save_as`, `timeout`, `retry_interval` |
| `click` | Click an element | `ref` (element ref or alias from `save_as`) |
| `right_click` | Right-click an element | `ref` |
| `type` | Type text into focused element | `text` |
| `send_keys` | Send keyboard shortcuts | `keys` |
| `wait` | Pause execution | `seconds` |
| `snapshot` | Capture the UI tree | `max_depth` |
| `screenshot` | Capture window image | — |
| `children` | List children of element | `ref`, `save_as` |
| `properties` | Get element properties | `ref` |
| `macro` | Execute a nested macro | `macro_name`, `params` |

### Parameter Substitution

Use `{{paramName}}` placeholders in any string property. Parameters are passed at invocation time via the `wpf_macro` MCP tool or the `macro` CLI command:

```
# MCP tool call
wpf_macro  name="acumen-fuse/import-file"  params={"filePath": "C:\\data\\schedule.xer"}

# CLI mode
macro acumen-fuse/import-file filePath=C:\data\schedule.xer
```

### Timeouts and Retry

- **Macro-level timeout** (`timeout` at root): Maximum seconds for the entire macro. Default: 60.
- **Step-level timeout** (`timeout` on a step): Maximum seconds for that step. Default: 5. For `find` and `find_by_path` actions, the step retries at `retry_interval` (default: 1s) until the timeout expires.
- **Crash detection**: If the target process exits mid-macro, execution stops immediately with a clear error.

### Examples

**Simple: Open File Menu**
```yaml
name: Open File Menu
description: Opens the Acumen Fuse File menu via ribbon keytips
timeout: 10
steps:
  - action: focus
  - action: send_keys
    keys: "Alt,F"
  - action: wait
    seconds: 0.5
```

**With Parameters: Import File**
```yaml
name: Import File
description: Opens the Import dialog and types a file path
timeout: 30
parameters:
  - name: filePath
    description: Full path to the file to import
    required: true
steps:
  - action: focus
  - action: send_keys
    keys: "Alt,F"
  - action: wait
    seconds: 0.5
  - action: find
    name: Import
    control_type: MenuItem
    save_as: importMenuItem
    timeout: 5
  - action: click
    ref: importMenuItem
  - action: wait
    seconds: 1
  - action: type
    text: "{{filePath}}"
  - action: send_keys
    keys: Enter
```

## Project Structure

```
C:\WpfMcp\
  WpfMcp.slnx                          Solution file
  .editorconfig                         Formatting standards (CRLF, 4-space indent, naming)
  nuget.config                          NuGet source configuration
  WpfMcp/
    WpfMcp.csproj                       Main project (.NET 9, WPF)
    Program.cs                          Entry point, mode routing, auto-launch logic
    Constants.cs                        Shared constants (pipe name, timeouts, JSON options)
    ElementCache.cs                     Thread-safe element reference cache (e1, e2, ...)
    Tools.cs                            MCP tool definitions (proxied and direct modes)
    UiaEngine.cs                        Core UI Automation engine (STA thread, SendInput)
    UiaProxy.cs                         Proxy client/server for named pipe communication
    MacroDefinition.cs                  YAML-deserialized POCOs for macros
    MacroEngine.cs                      Load, validate, and execute macro YAML files
    CliMode.cs                          Interactive CLI for manual testing
  macros/
    acumen-fuse/
      open-file-menu.yaml              Focus + Alt,F (example)
      find-projects-view.yaml          Find element with retry (example)
      import-file.yaml                 Full workflow with parameters (example)
  WpfMcp.Tests/
    WpfMcp.Tests.csproj                 xUnit test project (46 tests)
    MacroEngineTests.cs                 18 tests for macro loading, validation, execution
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
