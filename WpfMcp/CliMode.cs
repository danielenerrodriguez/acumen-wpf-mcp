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

    public static async Task RunAsync(string[] args)
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
        Console.WriteLine("  screenshot                 - Take screenshot (saves to wpfmcp_screenshot.png)");
        Console.WriteLine("  status                     - Show attachment status");
        Console.WriteLine("  quit                       - Exit");
        Console.WriteLine();

        var engine = UiaEngine.Instance;
        var cache = new ElementCache();

        while (true)
        {
            Console.Write("wpf> ");
            var line = Console.ReadLine();
            if (line == null) break;
            line = line.Trim();
            if (string.IsNullOrEmpty(line)) continue;

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
}
