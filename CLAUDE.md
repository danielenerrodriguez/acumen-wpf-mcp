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

# Test (126 tests)
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
`launch`, `wait_for_window`, `attach`, `focus`, `find`, `find_by_path`, `click`, `right_click`, `type`, `set_value`, `get_value`, `send_keys`, `keys` (alias), `wait`, `snapshot`, `screenshot`, `properties`, `children`, `file_dialog`, `macro`

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

## Knowledge Base System

Product-specific knowledge bases provide AI agents with the context they need to navigate WPF applications — automation IDs, keytips, keyboard shortcuts, ribbon structure, workflows, and more.

### File Convention
- Knowledge bases are YAML files named `_knowledge.yaml` (underscore prefix)
- Located in product subfolders: `publish/macros/{product}/_knowledge.yaml`
- Skipped by the macro loader (underscore prefix exclusion)
- Must have `kind: knowledge-base` as a top-level field

### How It Works
1. **`MacroEngine.Reload()`** loads `_knowledge.yaml` files into `_knowledgeBases` dictionary (keyed by product folder name)
2. **`wpf_macro_list`** includes a condensed summary for each knowledge base (verified keytips, key automation IDs, workflow/tip counts)
3. **MCP Resource** `knowledge://{productName}` serves the full YAML content via `Resources.cs`
4. Parsed as `Dictionary<string, object>` (not POCOs) — flexible, no C# changes needed when YAML schema changes

### Knowledge Base YAML Structure
```yaml
kind: knowledge-base
update_instructions: { ... }    # When/how AI agents should update this file
installation: { ... }           # Exe path, samples, templates, skills directories
application: { ... }            # Name, version, startup phases
keyboard_shortcuts: [ ... ]     # InputBinding shortcuts (Ctrl+N, Ctrl+O, etc.)
keytips: { ... }                # Application menu + ribbon tab keytip sequences
ribbon: { ... }                 # Complete ribbon structure (tabs, groups, tools, commands)
automation_ids: { ... }         # 80+ IDs by category (panels, grids, trees, etc.)
workflows: [ ... ]              # Step-by-step recipes for common tasks
navigation_tips: [ ... ]        # Practical tips for driving the application
data_formats: { ... }           # Supported import/export formats
```

### Current Knowledge Bases
- `acumen-fuse` — Deltek Acumen Fuse (1800+ lines, 100+ automation IDs, 15 workflows, 19 navigation tips)

## File Reference

| File | Purpose |
|------|---------|
| `WpfMcp/Program.cs` | Entry point: mode routing (`--server`, `--mcp-connect`, `--mcp`, drag-drop, CLI) |
| `WpfMcp/Constants.cs` | Shared constants, `ResolveMacrosPath()` |
| `WpfMcp/Tools.cs` | 19 MCP tool definitions (16 core + 3 recording) |
| `WpfMcp/UiaEngine.cs` | Core UI Automation engine (STA thread, SendInput, launch, wait) |
| `WpfMcp/UiaProxy.cs` | Proxy client/server over named pipe |
| `WpfMcp/MacroDefinition.cs` | YAML POCOs for macros + `KnowledgeBase` + `SaveMacroResult` records |
| `WpfMcp/MacroEngine.cs` | Load/validate/execute/save macros, FileSystemWatcher, knowledge base loading |
| `WpfMcp/MacroSerializer.cs` | YAML serialization, `BuildFromRecordedActions` |
| `WpfMcp/InputRecorder.cs` | Low-level hooks for recording user interactions |
| `WpfMcp/CliMode.cs` | Interactive CLI for manual testing |
| `WpfMcp/ElementCache.cs` | Thread-safe element reference cache (e1, e2, ...) |
| `WpfMcp/Resources.cs` | MCP resources — `knowledge://{productName}` endpoint |

## Macro Saving (`wpf_save_macro`)

AI agents can save workflows they've performed as reusable macro YAML files using the `wpf_save_macro` MCP tool.

### How It Works
1. Agent passes steps as a JSON array, plus name/description/parameters
2. `MacroEngine.ValidateSteps()` checks all action types and required fields
3. `MacroEngine.GetProductFolder()` scans knowledge bases for a matching `process_name` field to auto-derive the product folder (e.g., `Fuse` → `acumen-fuse`)
4. Steps are serialized to clean YAML (no JSON-in-YAML) via YamlDotNet
5. File is written to `{macrosPath}/{productFolder}/{name}.yaml`
6. FileSystemWatcher auto-reloads the new macro immediately

### Key Design Decisions
- Steps passed as JSON string in, written as clean YAML out
- Product folder auto-derived from attached process matching knowledge base `application.process_name`
- Validation against 20 known action types with per-action required field checks
- `force` parameter (bool, default false) for overwrite protection
- `SaveMacroResult` record returns `(Ok, FilePath, MacroName, Message)`

### Error Cases
1. No process attached → error
2. Can't derive product folder → error with suggestion to include folder in name
3. Macro already exists and `force` is false → error with overwrite hint
4. Invalid step action → error listing all valid actions
5. Missing required step fields → error naming the missing field and step number

### Files Modified
- `MacroDefinition.cs` — `SaveMacroResult` record
- `MacroEngine.cs` — `SaveMacro()`, `GetProductFolder()`, `ValidateSteps()`, `MacrosPath` property
- `Tools.cs` — `wpf_save_macro` MCP tool + `ParseStepsJson()` / `ConvertJsonElement()` helpers
- `UiaProxy.cs` — `saveMacro` case in `ExecuteCommand()` + `ParseJsonArray()` / `ConvertJsonElement()` helpers
- `UiaEngine.cs` — `ProcessName` property
- `_knowledge.yaml` — `saving_workflows` section under `update_instructions`

## Discoveries & Gotchas

- **MCP Tool Timeout (~15-20s)**: Keep macros fast. Reduce `wait` steps to 2-3s max. Total cumulative wait across nested macros must stay under ~12-15s.
- **File dialogs**: Two types — standard Win32 (`AutomationId="1148"`, `wpf_file_dialog` works) and DirectUI Save (`AutomationId="FileNameControlHost"`, must use `wpf_find` + `wpf_type` + `Enter`).
- **`wpf_screenshot` only captures main app window** — modal OS dialogs are NOT visible.
- **Sample file gotcha**: `Initial  Plan.xer` has a double space in the filename.
