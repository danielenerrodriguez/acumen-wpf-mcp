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
│  Tools.cs           18 MCP tool definitions (incl. macros+rec)  │
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
3. Bundles `publish/macros/` alongside the exe
4. Creates a GitHub Release with `WpfMcp.zip` (exe + macros) attached

Pull requests to `master` trigger build and test only (no release).

Extract the zip and keep `WpfMcp.exe` and the `macros/` folder in the same directory. The server discovers macros from the `macros/` folder next to the exe by default.

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
| `wpf_macro_list` | List all available macros, parameters, and knowledge base summaries |
| `wpf_record_start` | Start recording user interactions as a macro |
| `wpf_record_stop` | Stop recording and save the macro YAML file |
| `wpf_record_status` | Check recording state, action count, and duration |

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

## Macro Recording

The macro recording feature captures user interactions with the target WPF application and generates YAML macro files automatically. This is useful for creating macro drafts that can then be reviewed and refined.

### How It Works

Recording uses **low-level Windows input hooks** (`WH_MOUSE_LL` and `WH_KEYBOARD_LL`) combined with UIA `AutomationElement.FromPoint()` to identify which element the user clicked:

1. **Mouse clicks** are filtered to only the target application's window. Each click resolves the element at the cursor position and records its AutomationId, Name, ClassName, and ControlType.
2. **Keyboard input** goes through a state machine that detects:
   - **Typing**: Consecutive character keystrokes are coalesced into a single `type` step (300ms window)
   - **Keyboard combos**: `Ctrl+S`, `Alt+Shift+N` (modifier keys held simultaneously)
   - **Sequential keys**: `Alt,F` (Alt released alone, then F pressed within 500ms — for ribbon keytips)
   - **Special keys**: `Enter`, `Tab`, `Escape`, F-keys, etc.
3. **Wait detection**: Gaps longer than 1.5 seconds between actions produce `wait` steps, capped at 10 seconds.
4. Each recorded click becomes a `find` + `click` step pair, using the best available element identifier (AutomationId > Name+ControlType > ClassName+ControlType).

### Recording Workflow

```
1. Attach to the target app     →  wpf_attach  process_name="Fuse"
2. Start recording               →  wpf_record_start  name="acumen-fuse/my-workflow"
3. Interact with the app         →  (click buttons, type text, use keyboard shortcuts)
4. Stop recording                →  wpf_record_stop
5. Review the generated YAML     →  AI reviews and edits the macro before use
6. Run the macro                 →  wpf_macro  name="acumen-fuse/my-workflow"
```

### CLI Commands

```
record-start <name>   Start recording (e.g., record-start acumen-fuse/my-workflow)
record-stop           Stop recording and save the YAML file
record-status         Show recording state, action count, and duration
```

### Output

The recorded macro is saved to the macros folder (next to the exe) using the name as a relative path:

- `record-start acumen-fuse/my-workflow` → `macros/acumen-fuse/my-workflow.yaml`
- `record-start quick-test` → `macros/quick-test.yaml`

Subdirectories are created automatically. The file watcher picks up the new file immediately.

The `wpf_record_stop` tool also returns the generated YAML in the response, so the AI agent can review and suggest edits before the macro is used in production.

### Important Notes

- Recording runs on the **elevated server process** (hooks require the same privilege level as the target)
- Only interactions with the **attached target process** are captured (mouse clicks filtered by window ownership, keyboard filtered by foreground window check)
- Mouse **moves** are ignored — only clicks are recorded
- The recorder must be started and stopped through the proxy (`--mcp-connect` mode) or CLI mode — not direct MCP mode

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
| `acumen-fuse` | `publish/macros/acumen-fuse/_knowledge.yaml` | 1400+ lines, 80+ automation IDs, 12 workflows, 14 navigation tips, verified keytip sequences |

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
    Constants.cs                        Shared constants (pipe name, timeouts, JSON options)
    ElementCache.cs                     Thread-safe element reference cache (e1, e2, ...)
    Tools.cs                            18 MCP tool definitions (15 core + 3 recording)
    UiaEngine.cs                        Core UI Automation engine (STA thread, SendInput)
    UiaProxy.cs                         Proxy client/server for named pipe communication
    MacroDefinition.cs                  YAML-deserialized POCOs for macros + KnowledgeBase record
    MacroEngine.cs                      Load, validate, execute macros; knowledge base loading
    MacroSerializer.cs                  YAML serialization and BuildFromRecordedActions
    InputRecorder.cs                    Low-level hooks, keyboard state machine, recording
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
    WpfMcp.exe                          Build output (gitignored)
    *.dll                               Build output (gitignored)
  WpfMcp.Tests/
    WpfMcp.Tests.csproj                 xUnit test project (94 tests)
    MacroEngineTests.cs                 27 tests for macro loading, validation, execution
    MacroSerializerTests.cs             16 tests for YAML serialization and action building
    InputRecorderTests.cs               13 tests for recorder state and wait computation
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
