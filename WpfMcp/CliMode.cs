using System.IO;

namespace WpfMcp;

/// <summary>
/// Interactive CLI mode for testing WPF MCP tools without the MCP protocol.
/// Run with: WpfMcp.exe --cli
/// </summary>
public static class CliMode
{
    // Macro cancellation: set when a macro is running, cancelled by Ctrl+C
    private static CancellationTokenSource? _macroCts;

    static void PrintResult((bool success, string message) result) =>
        Console.WriteLine(result.success ? $"OK: {result.message}" : $"Error: {result.message}");

    public static async Task RunAsync(string[] args, string? macrosPath = null)
    {
        Console.WriteLine("WPF UIA MCP - CLI Test Mode");
        Console.WriteLine("===========================");
        Console.WriteLine("Commands:");
        Console.WriteLine("  attach <process_name>     - Attach to a running process");
        Console.WriteLine("  attach-pid <pid>          - Attach by PID");
        Console.WriteLine("  snapshot [depth]           - Get UI tree snapshot (default depth 3)");
        Console.WriteLine("  children [ref]             - List children of element (or root)");
        Console.WriteLine("  find <prop>=<val> ...      - Find element (automationid=, name=, classname=, controltype=)");
        Console.WriteLine("  find-path <seg1> <seg2>    - Find by ObjectStore path segments");
        Console.WriteLine("  props <ref>                - Get element properties");
        Console.WriteLine("  click <ref>                - Click element");
        Console.WriteLine("  rclick <ref>               - Right-click element");
        Console.WriteLine("  focus                      - Focus the attached window");
        Console.WriteLine("  keys <keys>                - Send keys (+ for combo, comma for sequence)");
        Console.WriteLine("  type <text>                - Type text");
        Console.WriteLine("  set-value <ref> <value>    - Set element value via ValuePattern");
        Console.WriteLine("  get-value <ref>            - Get element value via ValuePattern");
        Console.WriteLine("  read-prop <ref> <property> - Read a property (value, name, toggle_state, is_enabled, expand_state, is_selected, control_type, automation_id)");
        Console.WriteLine("  verify <ref> <prop> <exp>  - Verify element property equals expected value");
        Console.WriteLine("  file-dialog <path>         - Navigate file dialog to select a file");
        Console.WriteLine("  screenshot                 - Take screenshot (saves to wpfmcp_screenshot.png)");
        Console.WriteLine("  status                     - Show attachment status");
        Console.WriteLine("  macros                     - List available macros");
        Console.WriteLine("  macro <name> [k=v ...]     - Run a macro with optional parameters");
        Console.WriteLine("  run <path.yaml> [k=v ...]  - Run a YAML macro file directly (or drag file here)");
        Console.WriteLine("  export <name>              - Export a macro as a Windows shortcut (.lnk)");
        Console.WriteLine("  export-all                 - Export all macros as Windows shortcuts");
        Console.WriteLine("  watch                      - Start a watch session (records focus/hover/property changes)");
        Console.WriteLine("  watch stop                 - Stop the watch session and print entries");
        Console.WriteLine("  watch status               - Print current/last watch session entries");
        Console.WriteLine("  quit                       - Exit");
        Console.WriteLine();

        var engine = UiaEngine.Instance;
        var cache = new ElementCache();
        var macroEngine = new MacroEngine(macrosPath);
        var resolvedMacrosPath = Constants.ResolveMacrosPath(macrosPath);
        var commandLock = new SemaphoreSlim(1, 1);
        var appState = new AppState(engine, cache, new Lazy<MacroEngine>(() => macroEngine), commandLock);

        // Ctrl+C handler: cancel running macro instead of killing the process
        Console.CancelKeyPress += (_, e) =>
        {
            var cts = _macroCts;
            if (cts != null)
            {
                e.Cancel = true; // Don't terminate the process
                try { cts.Cancel(); } catch (ObjectDisposedException) { }
            }
            // If no macro running, let default Ctrl+C behavior terminate
        };

        while (true)
        {
            Console.Write("wpf> ");
            var line = Console.ReadLine();
            if (line == null) break;
            line = line.Trim();
            if (string.IsNullOrEmpty(line)) continue;

            // Detect bare .yaml path (typed, pasted, or from drag-drop onto console)
            var trimmedLine = line.Trim('"'); // Windows sometimes wraps dragged paths in quotes
            if (trimmedLine.EndsWith(".yaml", StringComparison.OrdinalIgnoreCase) && File.Exists(trimmedLine))
            {
                using var cts = new CancellationTokenSource();
                _macroCts = cts;
                try
                {
                    await RunYamlFileAsync(trimmedLine, "", macroEngine, engine, cache, cts.Token);
                }
                finally
                {
                    _macroCts = null;
                }
                continue;
            }

            var parts = line.Split(' ', 2, StringSplitOptions.TrimEntries);
            var cmd = parts[0].ToLower();
            var arg = parts.Length > 1 ? parts[1] : "";

            try
            {
                switch (cmd)
                {
                    case "quit" or "exit" or "q":
                        Console.WriteLine("Bye.");
                        return;

                    case "attach":
                        if (string.IsNullOrEmpty(arg)) { Console.WriteLine("Usage: attach <process_name>"); break; }
                        PrintResult(engine.Attach(arg));
                        break;

                    case "attach-pid":
                        if (!int.TryParse(arg, out int pid)) { Console.WriteLine("Usage: attach-pid <pid>"); break; }
                        PrintResult(engine.AttachByPid(pid));
                        break;

                    case "snapshot":
                        int depth = 3;
                        if (!string.IsNullOrEmpty(arg)) int.TryParse(arg, out depth);
                        var snapResult = engine.GetSnapshot(depth);
                        Console.WriteLine(snapResult.success ? snapResult.tree : $"Error: {snapResult.tree}");
                        break;

                    case "children":
                        System.Windows.Automation.AutomationElement? parent = null;
                        if (!string.IsNullOrEmpty(arg))
                        {
                            if (!cache.TryGet(arg, out parent))
                            { Console.WriteLine($"Unknown ref '{arg}'"); break; }
                        }
                        var children = engine.GetChildElements(parent);
                        if (children.Count == 0) { Console.WriteLine("No children found"); break; }
                        Console.WriteLine($"Found {children.Count} children:");
                        foreach (var child in children)
                        {
                            var refKey = cache.Add(child);
                            Console.WriteLine($"  [{refKey}] {engine.FormatElement(child)}");
                        }
                        break;

                    case "find":
                        if (string.IsNullOrEmpty(arg)) { Console.WriteLine("Usage: find automationid=X name=Y classname=Z controltype=W"); break; }
                        string? fAutomationId = null, fName = null, fClassName = null, fControlType = null;
                        foreach (var pair in arg.Split(' ', StringSplitOptions.RemoveEmptyEntries))
                        {
                            var kv = pair.Split('=', 2);
                            if (kv.Length != 2) continue;
                            switch (kv[0].ToLower())
                            {
                                case "automationid": fAutomationId = kv[1]; break;
                                case "name": fName = kv[1]; break;
                                case "classname": fClassName = kv[1]; break;
                                case "controltype": fControlType = kv[1]; break;
                            }
                        }
                        var findResult = engine.FindElement(fAutomationId, fName, fClassName, fControlType);
                        if (findResult.success && findResult.element != null)
                        {
                            var refKey = cache.Add(findResult.element);
                            Console.WriteLine($"OK [{refKey}]: {findResult.message}");
                            var props = engine.GetElementProperties(findResult.element);
                            foreach (var kv in props) Console.WriteLine($"  {kv.Key}: {kv.Value}");
                        }
                        else Console.WriteLine($"Not found: {findResult.message}");
                        break;

                    case "find-path":
                        if (string.IsNullOrEmpty(arg)) { Console.WriteLine("Usage: find-path <seg1> <seg2> ..."); break; }
                        var segments = arg.Split(' ', StringSplitOptions.RemoveEmptyEntries).ToList();
                        var pathResult = engine.FindElementByPath(segments);
                        if (pathResult.success && pathResult.element != null)
                        {
                            var refKey = cache.Add(pathResult.element);
                            Console.WriteLine($"OK [{refKey}]: {pathResult.message}");
                            var props = engine.GetElementProperties(pathResult.element);
                            foreach (var kv in props) Console.WriteLine($"  {kv.Key}: {kv.Value}");
                        }
                        else Console.WriteLine($"Not found: {pathResult.message}");
                        break;

                    case "props":
                        if (string.IsNullOrEmpty(arg) || !cache.TryGet(arg, out var propsEl))
                        { Console.WriteLine($"Unknown ref '{arg}'"); break; }
                        var elProps = engine.GetElementProperties(propsEl!);
                        foreach (var kv in elProps) Console.WriteLine($"  {kv.Key}: {kv.Value}");
                        break;

                    case "click":
                        if (string.IsNullOrEmpty(arg) || !cache.TryGet(arg, out var clickEl))
                        { Console.WriteLine($"Unknown ref '{arg}'"); break; }
                        PrintResult(engine.ClickElement(clickEl!));
                        break;

                    case "rclick":
                        if (string.IsNullOrEmpty(arg) || !cache.TryGet(arg, out var rclickEl))
                        { Console.WriteLine($"Unknown ref '{arg}'"); break; }
                        PrintResult(engine.RightClickElement(rclickEl!));
                        break;

                    case "focus":
                        PrintResult(engine.FocusWindow());
                        break;

                    case "keys":
                        if (string.IsNullOrEmpty(arg)) { Console.WriteLine("Usage: keys Alt,F  or  keys Ctrl+S"); break; }
                        PrintResult(engine.SendKeyboardShortcut(arg));
                        break;

                    case "type":
                        if (string.IsNullOrEmpty(arg)) { Console.WriteLine("Usage: type <text>"); break; }
                        PrintResult(engine.TypeText(arg));
                        break;

                    case "set-value":
                    {
                        var svParts = arg?.Split(' ', 2);
                        if (svParts == null || svParts.Length < 2) { Console.WriteLine("Usage: set-value <ref> <value>"); break; }
                        if (!cache.TryGet(svParts[0], out var svEl)) { Console.WriteLine($"Error: Unknown ref '{svParts[0]}'"); break; }
                        PrintResult(engine.SetElementValue(svEl!, svParts[1]));
                        break;
                    }

                    case "get-value":
                    {
                        if (string.IsNullOrEmpty(arg)) { Console.WriteLine("Usage: get-value <ref>"); break; }
                        if (!cache.TryGet(arg, out var gvEl)) { Console.WriteLine($"Error: Unknown ref '{arg}'"); break; }
                        var gvResult = engine.GetElementValue(gvEl!);
                        Console.WriteLine(gvResult.success ? $"OK: {gvResult.message}" : $"Error: {gvResult.message}");
                        break;
                    }

                    case "read-prop":
                    {
                        var rpParts = arg?.Split(' ', 2, StringSplitOptions.TrimEntries);
                        if (rpParts == null || rpParts.Length < 2) { Console.WriteLine("Usage: read-prop <ref> <property>"); break; }
                        if (!cache.TryGet(rpParts[0], out var rpEl)) { Console.WriteLine($"Error: Unknown ref '{rpParts[0]}'"); break; }
                        var rpResult = engine.ReadElementProperty(rpEl!, rpParts[1]);
                        Console.WriteLine(rpResult.success ? $"OK: {rpResult.message}" : $"Error: {rpResult.message}");
                        break;
                    }

                    case "verify":
                    {
                        var vParts = arg?.Split(' ', 4, StringSplitOptions.TrimEntries);
                        if (vParts == null || vParts.Length < 3)
                        {
                            Console.WriteLine("Usage: verify <ref> <property> <expected> [match_mode]");
                            Console.WriteLine("  Properties: value, name, toggle_state, is_enabled, expand_state, is_selected, control_type, automation_id");
                            Console.WriteLine("  Match modes: equals (default), contains, not_equals, regex, starts_with");
                            break;
                        }
                        if (!cache.TryGet(vParts[0], out var vEl)) { Console.WriteLine($"Error: Unknown ref '{vParts[0]}'"); break; }
                        var vResult = engine.ReadElementProperty(vEl!, vParts[1]);
                        if (!vResult.success) { Console.WriteLine($"Error: {vResult.message}"); break; }
                        var vMatchMode = vParts.Length >= 4 ? vParts[3].ToLowerInvariant() : "equals";
                        var vMatched = MacroEngine.VerifyMatch(vResult.value ?? "", vParts[2], vMatchMode);
                        if (vMatched == null) { Console.WriteLine($"Error: Unknown match_mode '{vParts[3]}'. Valid: equals, contains, not_equals, regex, starts_with"); break; }
                        if (vMatched.Value)
                            Console.WriteLine($"PASS ({vMatchMode}): {vParts[1]} = \"{vResult.value}\"");
                        else
                            Console.WriteLine($"FAIL ({vMatchMode}): expected {vParts[1]} = \"{vParts[2]}\" but got \"{vResult.value}\"");
                        break;
                    }

                    case "file-dialog":
                        if (string.IsNullOrEmpty(arg)) { Console.WriteLine("Usage: file-dialog <full-file-path>"); break; }
                        var fdResult = engine.FileDialogSetPath(arg);
                        Console.WriteLine(fdResult.success ? $"OK: {fdResult.message}" : $"Error: {fdResult.message}");
                        break;

                    case "screenshot":
                        var ssResult = engine.TakeScreenshot();
                        if (ssResult.success)
                        {
                            var bytes = Convert.FromBase64String(ssResult.base64);
                            var path = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), "wpfmcp_screenshot.png");
                            File.WriteAllBytes(path, bytes);
                            Console.WriteLine($"OK: {ssResult.message} -> saved to {path}");
                        }
                        else Console.WriteLine($"Error: {ssResult.message}");
                        break;

                    case "status":
                        Console.WriteLine($"Attached: {engine.IsAttached}");
                        if (engine.IsAttached)
                        {
                            Console.WriteLine($"  Window: {engine.WindowTitle}");
                            Console.WriteLine($"  PID: {engine.ProcessId}");
                        }
                        Console.WriteLine($"  Cached elements: {cache.Count}");
                        break;

                    case "macros":
                        var macroList = macroEngine.List();
                        var loadErrors = macroEngine.LoadErrors;
                        var knowledgeBases = macroEngine.KnowledgeBases;
                        if (macroList.Count == 0 && loadErrors.Count == 0 && knowledgeBases.Count == 0)
                        {
                            Console.WriteLine("No macros found. Place .yaml files in the macros/ folder.");
                            break;
                        }
                        if (macroList.Count > 0)
                        {
                            Console.WriteLine($"Available macros ({macroList.Count}):");
                            foreach (var m in macroList)
                            {
                                Console.WriteLine($"  {m.Name} - {m.Description}");
                                foreach (var p in m.Parameters)
                                    Console.WriteLine($"    {p.Name}{(p.Required ? " (required)" : "")} - {p.Description}{(p.Default != null ? $" [default: {p.Default}]" : "")}");
                            }
                        }
                        if (knowledgeBases.Count > 0)
                        {
                            Console.WriteLine($"\nKnowledge bases ({knowledgeBases.Count}):");
                            foreach (var kb in knowledgeBases)
                                Console.WriteLine(kb.Summary);
                        }
                        if (loadErrors.Count > 0)
                        {
                            Console.WriteLine($"\nLoad errors ({loadErrors.Count}):");
                            foreach (var err in loadErrors)
                                Console.WriteLine($"  {err.FilePath}: {err.Error}");
                        }
                        break;

                    case "macro":
                        if (string.IsNullOrEmpty(arg))
                        {
                            Console.WriteLine("Usage: macro <name> [param1=value1 param2=value2 ...]");
                            break;
                        }
                        var macroParts = arg.Split(' ', 2, StringSplitOptions.TrimEntries);
                        var macroName = macroParts[0];
                        var macroParams = new Dictionary<string, string>();
                        if (macroParts.Length > 1)
                        {
                            foreach (var pair in macroParts[1].Split(' ', StringSplitOptions.RemoveEmptyEntries))
                            {
                                var kv = pair.Split('=', 2);
                                if (kv.Length == 2) macroParams[kv[0]] = kv[1];
                            }
                        }
                        Console.WriteLine($"Running macro '{macroName}'... (Ctrl+C to cancel)");
                        using (var cts = new CancellationTokenSource())
                        {
                            _macroCts = cts;
                            try
                            {
                                var macroResult = await macroEngine.ExecuteAsync(macroName, macroParams, engine, cache,
                                    cancellation: cts.Token);
                                if (macroResult.Success)
                                    Console.WriteLine($"OK: {macroResult.Message}");
                                else
                                {
                                    Console.WriteLine($"FAILED: {macroResult.Message}");
                                    if (macroResult.Error != null)
                                        Console.WriteLine($"  Error: {macroResult.Error}");
                                    Console.WriteLine($"  Steps completed: {macroResult.StepsExecuted}/{macroResult.TotalSteps}");
                                }
                            }
                            finally
                            {
                                _macroCts = null;
                            }
                        }
                        break;

                    case "run":
                        if (string.IsNullOrEmpty(arg))
                        {
                            Console.WriteLine("Usage: run <path.yaml> [param1=value1 param2=value2 ...]");
                            break;
                        }
                        // Split: first token is the file path, rest are params
                        var runParts = arg.Split(' ', 2, StringSplitOptions.TrimEntries);
                        var runPath = runParts[0].Trim('"');
                        var runParamsStr = runParts.Length > 1 ? runParts[1] : "";
                        using (var runCts = new CancellationTokenSource())
                        {
                            _macroCts = runCts;
                            try
                            {
                                await RunYamlFileAsync(runPath, runParamsStr, macroEngine, engine, cache, runCts.Token);
                            }
                            finally
                            {
                                _macroCts = null;
                            }
                        }
                        break;

                    case "export":
                        if (string.IsNullOrEmpty(arg))
                        {
                            Console.WriteLine("Usage: export <macro-name>");
                            Console.WriteLine("  Exports a macro as a Windows shortcut (.lnk) to the Shortcuts/ folder.");
                            break;
                        }
                        var exportResult = macroEngine.ExportMacro(arg.Trim());
                        Console.WriteLine(exportResult.Ok
                            ? $"OK: {exportResult.Message}"
                            : $"Error: {exportResult.Message}");
                        break;

                    case "export-all":
                    {
                        var exportResults = macroEngine.ExportAllMacros();
                        int exportOk = 0, exportFail = 0;
                        foreach (var er in exportResults)
                        {
                            Console.WriteLine(er.Ok
                                ? $"  OK: {er.MacroName} -> {er.ShortcutPath}"
                                : $"  FAILED: {er.MacroName} - {er.Message}");
                            if (er.Ok) exportOk++; else exportFail++;
                        }
                        Console.WriteLine($"Exported {exportOk} shortcut(s), {exportFail} failed.");
                        break;
                    }

                    case "watch":
                    {
                        var sub = arg?.Trim().ToLowerInvariant();
                        if (sub == "stop")
                        {
                            var session = appState.StopWatchAsync().GetAwaiter().GetResult();
                            if (session == null)
                            {
                                Console.WriteLine("No watch session active.");
                                break;
                            }
                            PrintWatchSession(session);
                        }
                        else if (sub == "status" || sub == "log")
                        {
                            var session = appState.GetWatchSession();
                            if (session == null)
                            {
                                Console.WriteLine("No watch session found.");
                                break;
                            }
                            PrintWatchSession(session);
                        }
                        else
                        {
                            var session = appState.StartWatchAsync().GetAwaiter().GetResult();
                            if (session == null)
                            {
                                Console.WriteLine("Watch session already active. Use 'watch stop' to stop it.");
                                break;
                            }
                            Console.WriteLine($"Watch started (session {session.Id}). Focus, hover, and property changes are being recorded.");
                            Console.WriteLine("Use 'watch stop' to stop and view entries, or 'watch status' to check progress.");
                        }
                        break;
                    }

                    default:
                        Console.WriteLine($"Unknown command: {cmd}. Type 'quit' to exit.");
                        break;
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Exception: {ex.Message}");
            }
        }
    }

    private static void PrintWatchSession(WpfMcp.Web.WatchSession session)
    {
        var status = session.IsActive ? "active" : session.StopTime?.ToString("HH:mm:ss") ?? "?";
        Console.WriteLine($"Session: {session.Id} | Started: {session.StartTime:HH:mm:ss} | Stopped: {status} | Entries: {session.Entries.Count}");
        Console.WriteLine();

        foreach (var e in session.Entries)
        {
            var elementDesc = !string.IsNullOrEmpty(e.AutomationId) ? $"[{e.ControlType}] #{e.AutomationId}"
                : !string.IsNullOrEmpty(e.Name) ? $"[{e.ControlType}] \"{e.Name}\""
                : $"[{e.ControlType}]";

            if (e.Kind == WpfMcp.Web.WatchEntryKind.PropertyChange)
            {
                Console.WriteLine($"  {e.Time:HH:mm:ss.fff} [PropChange] {elementDesc}  {e.ChangedProperty}: \"{e.OldValue}\" → \"{e.NewValue}\"");
            }
            else if (e.Kind == WpfMcp.Web.WatchEntryKind.Keypress)
            {
                var context = !string.IsNullOrEmpty(e.AutomationId) || !string.IsNullOrEmpty(e.Name) || !string.IsNullOrEmpty(e.ControlType)
                    ? $"  (on {elementDesc})" : "";
                Console.WriteLine($"  {e.Time:HH:mm:ss.fff} [Keypress]  {e.KeyCombo}{context}");
            }
            else
            {
                var extras = new List<string>();
                if (e.Properties.TryGetValue("Value", out var val) && !string.IsNullOrEmpty(val))
                    extras.Add($"Value=\"{(val.Length > 40 ? val[..40] + "..." : val)}\"");
                if (e.Properties.TryGetValue("ToggleState", out var ts))
                    extras.Add($"Toggle={ts}");
                if (e.Properties.TryGetValue("IsSelected", out var sel) && sel == "True")
                    extras.Add("Selected");
                if (e.Properties.TryGetValue("ExpandCollapseState", out var ecs) && ecs != "LeafNode")
                    extras.Add(ecs);

                var suffix = extras.Count > 0 ? $"  ({string.Join(", ", extras)})" : "";
                Console.WriteLine($"  {e.Time:HH:mm:ss.fff} [{e.Kind}]  {elementDesc}{suffix}");
            }
        }
    }

    /// <summary>
    /// Load a YAML macro file, prompt for missing required parameters, and execute it.
    /// </summary>
    private static async Task RunYamlFileAsync(
        string yamlPath, string paramsStr,
        MacroEngine macroEngine, UiaEngine engine, ElementCache cache,
        CancellationToken cancellation = default)
    {
        if (!File.Exists(yamlPath))
        {
            Console.WriteLine($"File not found: {yamlPath}");
            return;
        }

        MacroDefinition macro;
        try
        {
            var yaml = File.ReadAllText(yamlPath);
            macro = YamlHelpers.Deserializer.Deserialize<MacroDefinition>(yaml);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Failed to parse YAML: {ex.Message}");
            return;
        }

        if (macro == null || macro.Steps.Count == 0)
        {
            Console.WriteLine("Invalid macro: no steps defined.");
            return;
        }

        // Parse any inline params
        var runParams = new Dictionary<string, string>();
        if (!string.IsNullOrEmpty(paramsStr))
        {
            foreach (var pair in paramsStr.Split(' ', StringSplitOptions.RemoveEmptyEntries))
            {
                var kv = pair.Split('=', 2);
                if (kv.Length == 2) runParams[kv[0]] = kv[1];
            }
        }

        // Prompt for missing required parameters
        foreach (var p in macro.Parameters)
        {
            if (p.Required && !runParams.ContainsKey(p.Name) && p.Default == null)
            {
                Console.Write($"  {p.Name}{(string.IsNullOrEmpty(p.Description) ? "" : $" ({p.Description})")}: ");
                var value = Console.ReadLine()?.Trim();
                if (!string.IsNullOrEmpty(value))
                    runParams[p.Name] = value;
            }
        }

        var displayName = macro.Name ?? Path.GetFileNameWithoutExtension(yamlPath);
        Console.WriteLine($"Running '{displayName}' ({macro.Steps.Count} steps)...");

        var result = await macroEngine.ExecuteDefinitionAsync(macro, displayName, runParams, engine, cache,
            cancellation: cancellation);
        if (result.Success)
            Console.WriteLine($"OK: {result.Message}");
        else
        {
            Console.WriteLine($"FAILED: {result.Message}");
            if (result.Error != null)
                Console.WriteLine($"  Error: {result.Error}");
            Console.WriteLine($"  Steps completed: {result.StepsExecuted}/{result.TotalSteps}");
        }
    }
}
