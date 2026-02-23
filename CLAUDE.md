# WPF MCP Server

Generic MCP (Model Context Protocol) server that automates any WPF application via UI Automation. Primary target is **Deltek Acumen Fuse** (`Fuse.exe`), which uses Infragistics controls.

## Project Structure

```
WpfMcp.slnx              # Solution (2 projects)
WpfMcp/                   # Main project (net9.0-windows, UseWPF)
WpfMcp.Tests/             # xUnit tests
publish/macros/           # Version-controlled macro YAML files (product subfolders)
.github/workflows/        # CI/CD: build, test, release zip
```

Flat layout — NO `src/` or `tests/` subdirectories.

## Architecture: Two-Process Proxy

```
OpenCode (WSL)
  → cmd.exe /c WpfMcp.exe --mcp-connect  (non-elevated, MCP stdio server)
    → auto-launches WpfMcp.exe --server   (elevated via BeyondTrust/UAC)
    → UiaProxyClient connects to WpfMcp_UIA named pipe
    → MCP tool calls proxy through pipe → UiaProxyServer executes UIA commands
```

**Why elevation?** Fuse's UIA providers only respond to elevated processes. Non-elevated processes see only `FrameworkId="Win32"` with a single TitleBar child.

## Build & Test

```bash
# MUST kill running instances first or build fails with file lock errors
cmd.exe /c "taskkill /IM WpfMcp.exe /F 2>nul"

# Build
cmd.exe /c "cd /d C:\WpfMcp && dotnet build WpfMcp.slnx"

# Test (94 tests)
cmd.exe /c "cd /d C:\WpfMcp && dotnet test WpfMcp.Tests"

# Release build + local publish
cmd.exe /c "cd /d C:\WpfMcp && dotnet build WpfMcp\WpfMcp.csproj -c Release"
```

The machine has .NET 10 preview SDK installed; we target `net9.0-windows`.

## Critical Technical Constraints

### WSL ↔ Windows
- OpenCode runs on WSL, launches Windows exe via `cmd.exe /c`
- WSL `/mnt/c/` writes can silently fail — use `cmd.exe` or `powershell.exe` for writing Windows files when reliability matters
- Use the `Write` tool for simple file writes; it works for most cases

### Named Pipe ACL
- Elevated server pipe needs `PipeSecurity` with `AuthenticatedUserSid` or non-elevated clients can't connect

### MCP Logging
- Must use `builder.Logging.ClearProviders()` then add console with `LogToStandardErrorThreshold = LogLevel.Trace` to prevent stdout corruption (MCP uses stdout for JSON-RPC)

### NuGet
- Local `nuget.config` with `<clear />` needed to avoid Infragistics NuGet source errors
- `ModelContextProtocol` NuGet version: `0.6.0-preview.1`

### Publish / Runtime
- `publish/` folder must include `runtimes/` subdirectory — `System.Text.Encodings.Web` v10.0.0.0 assembly is needed at runtime
- `Environment.ProcessPath` for exe location (NOT `AppContext.BaseDirectory` which points to temp extraction dir for single-file publish)

### Process Launch Relay
- Some apps (Windows 11 Notepad, store apps) exit immediately on `Process.Start` and relaunch under a different PID
- `WaitForProcessByNameAsync` handles this by falling back to process-name search when the tracked PID exits

## Key Code Patterns

### JsonElement Null Safety
When MCP passes `null` for optional params, always check `ValueKind` before calling `GetString()`/`GetInt32()`:
```csharp
private static string? GetStringArg(JsonElement args, string name) =>
    args.TryGetProperty(name, out var v) && v.ValueKind == JsonValueKind.String ? v.GetString() : null;
```

### YamlDotNet Configuration
Always use `UnderscoredNamingConvention` and `IgnoreUnmatchedProperties`. Use `[YamlMember(Alias = "...")]` attributes on POCOs.

### ControlType Lookup
Uses reflection-based static dictionary on `ControlType`'s public fields.

### Key Parsing
`ParseVirtualKey` uses `System.Windows.Input.Key` enum + `KeyInterop.VirtualKeyFromKey()`. Bare digits "0"-"9" must be aliased to "D0"-"D9".

### xUnit Test Isolation
- Pipe tests must use unique pipe names per test (GUID suffix) since xUnit runs in parallel
- MacroEngine tests use `enableWatcher: false` constructor param to avoid filesystem race conditions
- FileSystemWatcher tests need ~1500ms delay for debounce (500ms) + processing

## Macro System

### YAML Format
```yaml
name: My Macro
description: What it does
timeout: 60
parameters:
  - name: paramName
    description: What this param is
    required: true
    default: "optional default"
steps:
  - action: launch
    exe_path: "{{paramName}}"
    if_not_running: true
    timeout: 45
  - action: wait_for_window
    title_contains: "My App"
    timeout: 30
  - action: focus
  - action: find
    automation_id: myElement
    save_as: el
  - action: click
    ref: el
  - action: type
    text: "{{paramName}}"
  - action: send_keys
    keys: "Ctrl+S"
  - action: wait
    seconds: 2
  - action: macro
    macro_name: subfolder/other-macro
    params:
      key: value
```

### Step Types
`launch`, `wait_for_window`, `attach`, `focus`, `find`, `find_by_path`, `click`, `right_click`, `type`, `send_keys`, `keys` (alias), `wait`, `snapshot`, `screenshot`, `properties`, `children`, `macro`

### Macros Path Resolution
`Constants.ResolveMacrosPath()`: explicit `--macros-path` arg > `WPFMCP_MACROS_PATH` env var > `macros/` next to exe

### File Watcher
`MacroEngine` watches the macros folder with FileSystemWatcher (500ms debounce). Auto-reloads on create/modify/delete/rename. `LoadErrors` property tracks YAML parse failures.

## Conventions (DLTKEngineering)

- Repo naming: `[team]-[repo]`, lowercase+hyphens
- Branch prefixes: `feature/`, `bugfix/`, `release/`
- xUnit for tests
- File-scoped namespaces (`namespace WpfMcp;`)
- Private fields: `_camelCase`
- Constants: `PascalCase`

## Git

- Remote: `https://github.com/danielenerrodriguez/acumen-wpf-mcp`
- Branch: `master`
- Do NOT push — let the user push manually
- Do NOT create ZenDesk tickets

## File Reference

| File | Purpose |
|------|---------|
| `WpfMcp/Program.cs` | Entry point: mode routing (`--server`, `--mcp-connect`, `--mcp`, drag-drop, CLI) |
| `WpfMcp/Constants.cs` | Shared constants, `ResolveMacrosPath()` |
| `WpfMcp/Tools.cs` | 18 MCP tool definitions (15 core + 3 recording) |
| `WpfMcp/UiaEngine.cs` | Core UI Automation engine (STA thread, SendInput, launch, wait) |
| `WpfMcp/UiaProxy.cs` | Proxy client/server over named pipe |
| `WpfMcp/MacroDefinition.cs` | YAML POCOs for macros |
| `WpfMcp/MacroEngine.cs` | Load/validate/execute macros, FileSystemWatcher |
| `WpfMcp/MacroSerializer.cs` | YAML serialization, `BuildFromRecordedActions` |
| `WpfMcp/InputRecorder.cs` | Low-level hooks for recording user interactions |
| `WpfMcp/CliMode.cs` | Interactive CLI for manual testing |
| `WpfMcp/ElementCache.cs` | Thread-safe element reference cache (e1, e2, ...) |
