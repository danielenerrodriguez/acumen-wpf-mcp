using System.IO;

namespace WpfMcp;

/// <summary>
/// Interactive CLI mode for testing WPF MCP tools without the MCP protocol.
/// Run with: WpfMcp.exe --cli
/// </summary>
public static class CliMode
{
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
        Console.WriteLine("  file-dialog <path>         - Navigate file dialog to select a file");
        Console.WriteLine("  screenshot                 - Take screenshot (saves to wpfmcp_screenshot.png)");
        Console.WriteLine("  status                     - Show attachment status");
        Console.WriteLine("  macros                     - List available macros");
        Console.WriteLine("  macro <name> [k=v ...]     - Run a macro with optional parameters");
        Console.WriteLine("  run <path.yaml> [k=v ...]  - Run a YAML macro file directly (or drag file here)");
        Console.WriteLine("  record-start <name>        - Start recording a macro");
        Console.WriteLine("  record-stop                - Stop recording and save");
        Console.WriteLine("  record-status              - Show recording status");
        Console.WriteLine("  quit                       - Exit");
        Console.WriteLine();

        var engine = UiaEngine.Instance;
        var cache = new ElementCache();
        var macroEngine = new MacroEngine(macrosPath);
        var recorder = new InputRecorder(engine);
        var resolvedMacrosPath = Constants.ResolveMacrosPath(macrosPath);

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
                await RunYamlFileAsync(trimmedLine, "", macroEngine, engine, cache);
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
                        if (macroList.Count == 0 && loadErrors.Count == 0)
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
                        Console.WriteLine($"Running macro '{macroName}'...");
                        var macroResult = await macroEngine.ExecuteAsync(macroName, macroParams, engine, cache);
                        if (macroResult.Success)
                            Console.WriteLine($"OK: {macroResult.Message}");
                        else
                        {
                            Console.WriteLine($"FAILED: {macroResult.Message}");
                            if (macroResult.Error != null)
                                Console.WriteLine($"  Error: {macroResult.Error}");
                            Console.WriteLine($"  Steps completed: {macroResult.StepsExecuted}/{macroResult.TotalSteps}");
                        }
                        break;

                    case "record-start":
                        if (string.IsNullOrEmpty(arg))
                        {
                            Console.WriteLine("Usage: record-start <name> (e.g., acumen-fuse/my-workflow)");
                            break;
                        }
                        var startResult = recorder.StartRecording(arg.Trim(), resolvedMacrosPath);
                        Console.WriteLine(startResult.success ? $"OK: {startResult.message}" : $"Error: {startResult.message}");
                        break;

                    case "record-stop":
                        var stopResult = recorder.StopRecording();
                        if (stopResult.success)
                        {
                            Console.WriteLine($"OK: {stopResult.message}");
                            if (stopResult.filePath != null)
                                Console.WriteLine($"  File: {stopResult.filePath}");
                            if (stopResult.yaml != null)
                            {
                                Console.WriteLine("  --- Generated YAML ---");
                                Console.WriteLine(stopResult.yaml);
                                Console.WriteLine("  --- End YAML ---");
                            }
                        }
                        else
                        {
                            Console.WriteLine($"Error: {stopResult.message}");
                        }
                        break;

                    case "record-status":
                        Console.WriteLine($"State: {recorder.State}");
                        if (recorder.State == InputRecorder.RecorderState.Recording)
                        {
                            Console.WriteLine($"  Macro: {recorder.MacroName}");
                            Console.WriteLine($"  Actions: {recorder.ActionCount}");
                            Console.WriteLine($"  Duration: {recorder.Duration:mm\\:ss}");
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
                        await RunYamlFileAsync(runPath, runParamsStr, macroEngine, engine, cache);
                        break;

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

    /// <summary>
    /// Load a YAML macro file, prompt for missing required parameters, and execute it.
    /// </summary>
    private static async Task RunYamlFileAsync(
        string yamlPath, string paramsStr,
        MacroEngine macroEngine, UiaEngine engine, ElementCache cache)
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
            var deserializer = new YamlDotNet.Serialization.DeserializerBuilder()
                .WithNamingConvention(YamlDotNet.Serialization.NamingConventions.UnderscoredNamingConvention.Instance)
                .IgnoreUnmatchedProperties()
                .Build();
            macro = deserializer.Deserialize<MacroDefinition>(yaml);
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

        var result = await macroEngine.ExecuteDefinitionAsync(macro, displayName, runParams, engine, cache);
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
