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
- **Knowledge bases** provide AI agents with structured navigation context (automation IDs, keytips, workflows) per application

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
│  Tools.cs           17 MCP tool definitions                     │
│  UiaProxyClient     Sends JSON requests over named pipe         │
│  ElementCache       Thread-safe LRU ref cache (e1, e2, max 500) │
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
3. Bundles `publish/macros/` and `setup.cmd` alongside the exe
4. Creates a GitHub Release with `WpfMcp.zip` (exe + macros + setup script) attached

Pull requests to `master` trigger build and test only (no release).

Extract the zip and keep `WpfMcp.exe` and the `macros/` folder in the same directory. Run `setup.cmd` to generate Windows shortcuts for all macros — they'll appear in a `Shortcuts/` folder next to `macros/`. The server discovers macros from the `macros/` folder next to the exe by default.

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

### Export Shortcuts

```
WpfMcp.exe --export-all [--shortcuts-path C:\path] [--force]
```

Exports all macros as Windows shortcut (.lnk) files. Users can double-click a shortcut to run the macro. The first run triggers a UAC prompt for the elevated server; subsequent runs connect to the existing server silently.

A `setup.cmd` script is included in the release zip for convenience — run it after extraction to generate all shortcuts.

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
| `wpf_macro_list` | List all available macros, parameters, and knowledge base summaries |
| `wpf_save_macro` | Save a workflow as a reusable macro YAML file (auto-derives product folder) |
| `wpf_export_macro` | Export a macro (or all macros) as a double-clickable Windows shortcut (.lnk) |

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

Macros are reusable, YAML-defined step sequences that automate common UI workflows. They support parameters, timeouts, and retry logic for find operations.

### Folder Structure

```
publish/
  macros/
    acumen-fuse/
      open-file-menu.yaml
      find-projects-view.yaml
      import-file.yaml
    another-app/
      some-workflow.yaml
  WpfMcp.exe          ← build output (gitignored)
  *.dll                ← build output (gitignored)
```

The `publish/macros/` folder is version-controlled and serves as the source of truth. On build, macros are copied to the build output directory so the exe finds them via the default `macros/` path next to itself. The `publish/` folder also receives the exe and dependency DLLs on Release builds, making it a ready-to-run local deployment.

You can override the macros location with:

- The `--macros-path <path>` CLI argument
- The `WPFMCP_MACROS_PATH` environment variable

Macro names are derived from the relative path without extension, using forward slashes (e.g., `acumen-fuse/import-file`).

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
| `set_value` | Set element value via UIA ValuePattern | `ref`, `value` |
| `get_value` | Get element value via UIA ValuePattern | `ref` |
| `file_dialog` | Navigate a file dialog to select a file | `text` (the file path) |
| `wait` | Pause execution | `seconds` |
| `snapshot` | Capture the UI tree | `max_depth` |
| `screenshot` | Capture window image | — |
| `children` | List children of element | `ref`, `save_as` |
| `properties` | Get element properties | `ref` |
| `launch` | Launch an application | `exe_path`, `arguments`, `working_directory`, `if_not_running`, `timeout` |
| `wait_for_window` | Wait for window to be ready | `title_contains`, `timeout`, `retry_interval` |
| `wait_for_enabled` | Wait for element to become enabled/disabled | `automation_id`, `name`, `ref`, `enabled`, `timeout`, `retry_interval` |
| `macro` | Execute a nested macro (not recommended — will exceed MCP client timeout) | `macro_name`, `params` |

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

## Macro Saving (AI-Driven)

The `wpf_save_macro` tool lets AI agents save workflows they've just performed as reusable macro YAML files. This is the primary way agents create new macros — they perform a task step by step, then save it for future reuse.

### How It Works

1. The agent performs a workflow using individual MCP tools (`wpf_find`, `wpf_click`, `wpf_send_keys`, etc.)
2. The agent calls `wpf_save_macro` with the steps as a JSON array
3. The tool auto-derives the product folder by matching the attached process name against knowledge base `process_name` fields
4. Steps are validated against 21 known action types with per-action required field checks
5. The macro is written as clean YAML to `{macrosPath}/{productFolder}/{name}.yaml`
6. The FileSystemWatcher auto-reloads it immediately — it's available via `wpf_macro` right away

### Example

```
wpf_save_macro(
  name: "switch-to-diagnostics",
  description: "Switch to the Diagnostics tab via keytips",
  steps: '[
    {"action":"focus"},
    {"action":"send_keys","keys":"Alt,2"},
    {"action":"wait","seconds":1}
  ]'
)
```

This creates `macros/acumen-fuse/switch-to-diagnostics.yaml` (product folder derived from the attached Fuse process).

### Overwrite Protection

By default, saving fails if a macro with the same name already exists. Pass `force=true` to overwrite.

### Validation

Steps are validated before writing:
- All action types must be known (21 supported actions)
- Required fields are checked per action type (e.g., `send_keys` requires `keys`, `find` requires at least one search property)
- Clear error messages reference the step number and missing field

## Macro Export (Shortcuts)

Macros can be exported as Windows shortcut (.lnk) files that users can double-click to run without needing an MCP client or AI agent.

### How It Works

The shortcut targets `WpfMcp.exe "C:\path\to\macro.yaml"`, reusing the existing drag-and-drop execution mode. This mode:

1. Parses the YAML macro file
2. Prompts for any required parameters that have no default value
3. Connects to the elevated server (launching it with UAC if not already running)
4. Executes the macro and displays the result

Shortcuts are placed in a `Shortcuts/` folder as a sibling of the `macros/` folder, mirroring the macro directory structure:

```
publish/
  macros/
    acumen-fuse/
      import-xer.yaml
      export-to-excel.yaml
  Shortcuts/
    acumen-fuse/
      import-xer.lnk
      export-to-excel.lnk
  setup.cmd
  WpfMcp.exe
```

### Generating Shortcuts

```bash
# Export all macros
WpfMcp.exe --export-all

# Export a single macro (via MCP tool or CLI)
wpf_export_macro name="acumen-fuse/import-xer"

# CLI mode
export acumen-fuse/import-xer
export-all

# Override output directory
WpfMcp.exe --export-all --shortcuts-path C:\Users\me\Desktop\MacroShortcuts

# Force overwrite existing shortcuts
WpfMcp.exe --export-all --force
```

### Elevation

Shortcuts do **not** have the "Run as administrator" flag. The client process runs non-elevated and handles elevation internally — it only triggers a UAC prompt if the elevated server isn't already running. This means:

- First shortcut double-click: UAC prompt appears once
- Subsequent runs: Connects to the existing elevated server silently
- The elevated server has a 5-minute idle timeout, then shuts down until needed again

### Path Resolution

The shortcuts output folder is resolved in this priority:

1. Explicit `--shortcuts-path` argument
2. `WPFMCP_SHORTCUTS_PATH` environment variable
3. `Shortcuts/` as a sibling of the macros folder (default)

## Knowledge Bases

Knowledge bases are YAML files that give AI agents the context they need to navigate a specific WPF application without trial and error. Instead of blind exploration via snapshots, the agent starts with a structured map of automation IDs, keytip sequences, ribbon layout, workflows, and sample file paths.

### How It Works

1. Place a `_knowledge.yaml` file in a product subfolder (e.g., `publish/macros/acumen-fuse/_knowledge.yaml`)
2. The file must have `kind: knowledge-base` as a top-level field
3. On startup, `MacroEngine` loads knowledge bases into memory (parsed as flexible `Dictionary<string, object>`, no C# class updates needed when the YAML schema changes)
4. `wpf_macro_list` includes a condensed summary alongside macros — keytip count, key automation IDs, workflow/tip counts
5. The full content is available as an MCP Resource at `knowledge://{productName}` (e.g., `knowledge://acumen-fuse`)

The underscore prefix (`_knowledge.yaml`) ensures these files are skipped by the macro loader — they aren't macros, they're reference data.

### YAML Structure

```yaml
kind: knowledge-base
update_instructions:              # When/how AI agents should update this file
  when_to_update: [...]
  validation_rules: [...]

installation:                     # Exe path, samples, templates, skills directories
  exe_path: "C:\\Program Files\\..."
  samples_directory: "..."

application:                      # Name, version, startup phases
  name: Deltek Acumen Fuse
  startup_phases: [...]

keyboard_shortcuts:               # Global shortcuts from InputBindings
  - shortcut: Ctrl+N
    action: New Workbook

keytips:                          # Ribbon keytip sequences (verified against live UI)
  application_menu:
    trigger: "Alt,F"
    items: [...]
  tabs: [...]

ribbon:                           # Complete ribbon structure (tabs, groups, tools, commands)
  tabs:
    - name: Projects
      automation_id: uxProjects
      groups: [...]

automation_ids:                   # All known IDs organized by category
  main_panels: [...]
  grids_and_trees: [...]

workflows:                        # Step-by-step recipes for common tasks
  - name: Import XER and Run Analysis
    steps: [...]

navigation_tips:                  # Practical tips for driving the app
  - "Use Alt keytips for reliable ribbon navigation..."

data_formats:                     # Supported import/export formats
  import: [...]
  export: [...]
```

### Current Knowledge Bases

| Product | File | Content |
|---------|------|---------|
| `acumen-fuse` | `publish/macros/acumen-fuse/_knowledge.yaml` | 1800+ lines, 100+ automation IDs, 15 workflows, 19 navigation tips, verified keytip sequences |

### MCP Resource Access

AI agents can read the full knowledge base via the MCP resource protocol:

```
Resource URI: knowledge://acumen-fuse
```

This returns the complete YAML content. The `wpf_macro_list` tool response includes a summary so agents know a knowledge base exists without needing to fetch the full resource.

## Project Structure

```
C:\WpfMcp\
  WpfMcp.slnx                          Solution file
  .editorconfig                         Formatting standards (CRLF, 4-space indent, naming)
  nuget.config                          NuGet source configuration
  WpfMcp/
    WpfMcp.csproj                       Main project (.NET 9, WPF)
    Program.cs                          Entry point, mode routing, auto-launch logic
    Constants.cs                        Shared constants (pipe name, timeouts, JSON options, Commands)
    ElementCache.cs                     Thread-safe LRU element reference cache (e1, e2, ..., max 500)
    Tools.cs                            17 MCP tool definitions
    UiaEngine.cs                        Core UI Automation engine (STA thread, SendInput)
    UiaProxy.cs                         Proxy client/server for named pipe communication
    MacroDefinition.cs                  YAML-deserialized POCOs for macros + KnowledgeBase + SaveMacroResult + ExportMacroResult records
    MacroEngine.cs                      Load, validate, execute, save, export macros; knowledge base loading
    ShortcutCreator.cs                  COM-based Windows shortcut (.lnk) creation via WScript.Shell
    MacroSerializer.cs                  YAML serialization (`ToYaml`, `SaveToFile`)
    JsonHelpers.cs                      Shared JSON array parsing utilities
    YamlHelpers.cs                      Shared YAML deserializer/serializer instances
    Resources.cs                        MCP resources (knowledge://{productName} endpoint)
    CliMode.cs                          Interactive CLI for manual testing
  publish/
    macros/                             Version-controlled macro + knowledge base YAML files
      acumen-fuse/
        _knowledge.yaml                Knowledge base (automation IDs, keytips, workflows)
        launch.yaml                    Launch Fuse and wait for ready state
        import-xer.yaml               Import XER file via file dialog
        import-xer-and-analyze.yaml    Import XER and run full analysis
        new-workbook.yaml              Create new workbook
        open-workbook.yaml             Open existing workbook via file dialog
        save-workbook.yaml             Save current workbook
        export-to-excel.yaml           Export data to Excel
        switch-to-tab.yaml            Switch ribbon tab by keytip
        ...                            (+ more macros)
    Shortcuts/                          Generated .lnk shortcuts (gitignored, machine-specific)
      acumen-fuse/
        import-xer.lnk                Double-click to run import-xer macro
        ...
    setup.cmd                           Post-extraction script to generate shortcuts
    WpfMcp.exe                          Build output (gitignored)
    *.dll                               Build output (gitignored)
  WpfMcp.Tests/
    WpfMcp.Tests.csproj                 xUnit test project (125 tests)
    MacroEngineTests.cs                 73 tests for macro loading, validation, execution, saving
    MacroExportTests.cs                 18 tests for shortcut export and ShortcutCreator
    MacroSerializerTests.cs             6 tests for YAML serialization round-trips
    WpfToolsTests.cs                    9 tests for MCP tool input validation
    ProxyResponseFormattingTests.cs     9 tests for proxy response formatting
    UiaProxyProtocolTests.cs            7 tests for named pipe protocol
    UiaProxyClientTests.cs              3 tests for proxy client edge cases
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
