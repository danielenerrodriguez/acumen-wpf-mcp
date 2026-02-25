using System.ComponentModel;
using System.Text.Json;
using System.Windows.Automation;
using ModelContextProtocol.Server;

namespace WpfMcp;

[McpServerToolType]
public static class WpfTools
{
    /// <summary>
    /// When set, all tool calls are proxied to the elevated server.
    /// Set by --mcp-connect mode before starting the MCP server.
    /// </summary>
    public static UiaProxyClient? Proxy { get; set; }

    // Local element cache (only used in non-proxied / direct mode)
    private static readonly ElementCache _cache = new();

    [McpServerTool, Description("Attach to a running WPF application by process name or PID. Optionally launch the application if it's not running and wait for window readiness.")]
    public static async Task<string> wpf_attach(
        [Description("Process name (without .exe)")] string? process_name = null,
        [Description("Process ID")] int? pid = null,
        [Description("Path to .exe — if provided and the process isn't running, launch it first")] string? exe_path = null,
        [Description("Command-line arguments when launching")] string? arguments = null,
        [Description("Wait until the window title contains this string (readiness check)")] string? wait_for_title = null,
        [Description("Seconds to wait for launch and readiness (default 30)")] int? timeout = null)
    {
        if (process_name == null && pid == null && exe_path == null)
            return "Error: Either process_name, pid, or exe_path must be provided";

        if (Proxy != null)
        {
            // If exe_path is provided, use the launch proxy command
            if (exe_path != null)
            {
                var launchArgs = new Dictionary<string, object?>
                {
                    ["exePath"] = exe_path,
                    ["arguments"] = arguments,
                    ["ifNotRunning"] = true,
                    ["timeout"] = timeout ?? 30
                };
                var launchResp = await Proxy.CallAsync(Constants.Commands.Launch, launchArgs);
                if (launchResp.TryGetProperty("ok", out var lok) && !lok.GetBoolean())
                    return FormatResponse(launchResp);

                // If wait_for_title specified, wait for window readiness
                if (!string.IsNullOrEmpty(wait_for_title))
                {
                    var waitArgs = new Dictionary<string, object?>
                    {
                        ["titleContains"] = wait_for_title,
                        ["timeout"] = timeout ?? 30
                    };
                    var waitResp = await Proxy.CallAsync(Constants.Commands.WaitForWindow, waitArgs);
                    if (waitResp.TryGetProperty("ok", out var wok) && !wok.GetBoolean())
                        return FormatResponse(waitResp);
                }

                return FormatResponse(launchResp);
            }

            var args = new Dictionary<string, object?> { ["processName"] = process_name, ["pid"] = pid };
            var resp = await Proxy.CallAsync(Constants.Commands.Attach, args);

            // If wait_for_title specified, wait for window readiness after attach
            if (resp.TryGetProperty("ok", out var aok) && aok.GetBoolean() && !string.IsNullOrEmpty(wait_for_title))
            {
                var waitArgs = new Dictionary<string, object?>
                {
                    ["titleContains"] = wait_for_title,
                    ["timeout"] = timeout ?? 30
                };
                var waitResp = await Proxy.CallAsync("waitForWindow", waitArgs);
                if (waitResp.TryGetProperty("ok", out var wok2) && !wok2.GetBoolean())
                    return FormatResponse(waitResp);
            }

            return FormatResponse(resp);
        }

        var engine = UiaEngine.Instance;

        // If exe_path provided, use launch-and-attach
        if (exe_path != null)
        {
            var launchResult = await engine.LaunchAndAttachAsync(
                exe_path, arguments, ifNotRunning: true,
                timeoutSec: timeout ?? 30);
            if (!launchResult.success)
                return $"Error: {launchResult.message}";

            if (!string.IsNullOrEmpty(wait_for_title))
            {
                var waitResult = await engine.WaitForWindowReadyAsync(
                    titleContains: wait_for_title, timeoutSec: timeout ?? 30);
                if (!waitResult.success)
                    return $"Error: {waitResult.message}";
            }

            return $"OK: {launchResult.message}";
        }

        var result = pid.HasValue ? engine.AttachByPid(pid.Value) : engine.Attach(process_name!);
        if (!result.success)
            return $"Error: {result.message}";

        if (!string.IsNullOrEmpty(wait_for_title))
        {
            var waitResult = await engine.WaitForWindowReadyAsync(
                titleContains: wait_for_title, timeoutSec: timeout ?? 30);
            if (!waitResult.success)
                return $"Error: {waitResult.message}";
        }

        return $"OK: {result.message}";
    }

    [McpServerTool, Description("Get a snapshot of the UI automation tree. Shows element types, names, automation IDs, and class names in a hierarchical view.")]
    public static async Task<string> wpf_snapshot(
        [Description("Maximum tree depth to explore (default 3)")] int max_depth = 3)
    {
        if (Proxy != null)
        {
            var args = new Dictionary<string, object?> { ["maxDepth"] = max_depth };
            var resp = await Proxy.CallAsync(Constants.Commands.Snapshot, args);
            return FormatResponse(resp);
        }

        var engine = UiaEngine.Instance;
        var result = engine.GetSnapshot(max_depth);
        return result.success ? result.tree : $"Error: {result.tree}";
    }

    [McpServerTool, Description("Find a UI element by its properties. Searches all descendants of the main window.")]
    public static async Task<string> wpf_find(
        [Description("AutomationId of the element")] string? automation_id = null,
        [Description("Name of the element")] string? name = null,
        [Description("ClassName of the element")] string? class_name = null,
        [Description("ControlType (e.g., Button, Edit, Custom, Tree, ToolBar)")] string? control_type = null)
    {
        if (Proxy != null)
        {
            var args = new Dictionary<string, object?>
            {
                ["automationId"] = automation_id, ["name"] = name,
                ["className"] = class_name, ["controlType"] = control_type
            };
            var resp = await Proxy.CallAsync(Constants.Commands.Find, args);
            return FormatProxyFindResponse(resp);
        }

        var engine = UiaEngine.Instance;
        var result = engine.FindElement(automation_id, name, class_name, control_type);
        if (result.success && result.element != null)
        {
            var refKey = _cache.Add(result.element);
            var props = engine.GetElementProperties(result.element);
            return $"OK [{refKey}]: {result.message}\nProperties: {JsonSerializer.Serialize(props, Constants.IndentedJson)}";
        }
        return $"Error: {result.message}";
    }

    [McpServerTool, Description("Find a UI element by walking a hierarchical path (ObjectStore format). Each path segment uses the format 'SearchProp:PropertyName~Value' with semicolons separating multiple properties.")]
    public static async Task<string> wpf_find_by_path(
        [Description("Array of path segments, each in format 'SearchProp:ControlType~Custom;SearchProp:AutomationId~uxProjectsView'")] string[] path)
    {
        if (Proxy != null)
        {
            var args = new Dictionary<string, object?> { ["path"] = path };
            var resp = await Proxy.CallAsync(Constants.Commands.FindByPath, args);
            return FormatProxyFindResponse(resp);
        }

        var engine = UiaEngine.Instance;
        var result = engine.FindElementByPath(path.ToList());
        if (result.success && result.element != null)
        {
            var refKey = _cache.Add(result.element);
            var props = engine.GetElementProperties(result.element);
            return $"OK [{refKey}]: {result.message}\nProperties: {JsonSerializer.Serialize(props, Constants.IndentedJson)}";
        }
        return $"Error: {result.message}";
    }

    [McpServerTool, Description("Get the children of an element (or the main window if no ref provided). Useful for exploring the UI tree.")]
    public static async Task<string> wpf_children(
        [Description("Element reference from a previous find (e.g., 'e1'). Omit for main window children.")] string? ref_key = null)
    {
        if (Proxy != null)
        {
            var args = new Dictionary<string, object?> { ["refKey"] = ref_key };
            var resp = await Proxy.CallAsync(Constants.Commands.Children, args);
            if (resp.TryGetProperty("ok", out var ok) && ok.GetBoolean())
            {
                var items = resp.GetProperty("result");
                var lines = new List<string> { $"Found {items.GetArrayLength()} children:" };
                foreach (var item in items.EnumerateArray())
                    lines.Add($"  [{item.GetProperty("ref").GetString()}] {item.GetProperty("desc").GetString()}");
                return lines.Count > 1 ? string.Join("\n", lines) : "No children found";
            }
            return $"Error: {resp.GetProperty("error").GetString()}";
        }

        var engine = UiaEngine.Instance;
        AutomationElement? parent = null;
        if (ref_key != null)
        {
            if (!_cache.TryGet(ref_key, out parent))
                return $"Error: Unknown element reference '{ref_key}'";
        }
        try
        {
            var children = engine.GetChildElements(parent);
            if (children.Count == 0) return "No children found";
            var lines = new List<string> { $"Found {children.Count} children:" };
            foreach (var child in children)
            {
                var childRef = _cache.Add(child);
                lines.Add($"  [{childRef}] {engine.FormatElement(child)}");
            }
            return string.Join("\n", lines);
        }
        catch (Exception ex) { return $"Error: {ex.Message}"; }
    }

    [McpServerTool, Description("Click on an element by its reference key.")]
    public static async Task<string> wpf_click(
        [Description("Element reference from a previous find (e.g., 'e1')")] string ref_key)
    {
        if (Proxy != null)
            return FormatResponse(await Proxy.CallAsync(Constants.Commands.Click, new() { ["refKey"] = ref_key }));

        if (!_cache.TryGet(ref_key, out var element))
            return $"Error: Unknown element reference '{ref_key}'";
        var result = UiaEngine.Instance.ClickElement(element!);
        return result.success ? $"OK: {result.message}" : $"Error: {result.message}";
    }

    [McpServerTool, Description("Right-click on an element by its reference key.")]
    public static async Task<string> wpf_right_click(
        [Description("Element reference from a previous find (e.g., 'e1')")] string ref_key)
    {
        if (Proxy != null)
            return FormatResponse(await Proxy.CallAsync(Constants.Commands.RightClick, new() { ["refKey"] = ref_key }));

        if (!_cache.TryGet(ref_key, out var element))
            return $"Error: Unknown element reference '{ref_key}'";
        var result = UiaEngine.Instance.RightClickElement(element!);
        return result.success ? $"OK: {result.message}" : $"Error: {result.message}";
    }

    [McpServerTool, Description("Type text into the currently focused element using simulated keystrokes. For reliable edit-field interaction (especially in dialogs), prefer wpf_set_value instead.")]
    public static async Task<string> wpf_type(
        [Description("Text to type")] string text)
    {
        if (Proxy != null)
            return FormatResponse(await Proxy.CallAsync(Constants.Commands.Type, new() { ["text"] = text }));

        var result = UiaEngine.Instance.TypeText(text);
        return result.success ? $"OK: {result.message}" : $"Error: {result.message}";
    }

    [McpServerTool, Description("Set the value of an element directly using UIA ValuePattern. Works reliably with edit fields and combo boxes in dialogs without focus issues.")]
    public static async Task<string> wpf_set_value(
        [Description("Element reference from a previous find (e.g., 'e1')")] string ref_key,
        [Description("Value to set")] string value)
    {
        if (Proxy != null)
            return FormatResponse(await Proxy.CallAsync(Constants.Commands.SetValue, new() { ["refKey"] = ref_key, ["value"] = value }));

        if (!_cache.TryGet(ref_key, out var element))
            return $"Error: Unknown element reference '{ref_key}'";
        var result = UiaEngine.Instance.SetElementValue(element!, value);
        return result.success ? $"OK: {result.message}" : $"Error: {result.message}";
    }

    [McpServerTool, Description("Get the current value of an element using UIA ValuePattern.")]
    public static async Task<string> wpf_get_value(
        [Description("Element reference from a previous find (e.g., 'e1')")] string ref_key)
    {
        if (Proxy != null)
        {
            var resp = await Proxy.CallAsync(Constants.Commands.GetValue, new() { ["refKey"] = ref_key });
            if (resp.TryGetProperty("ok", out var ok) && ok.GetBoolean())
                return $"OK: {resp.GetProperty("result").GetString()}";
            return $"Error: {resp.GetProperty("error").GetString()}";
        }

        if (!_cache.TryGet(ref_key, out var element))
            return $"Error: Unknown element reference '{ref_key}'";
        var result = UiaEngine.Instance.GetElementValue(element!);
        return result.success ? $"OK: {result.message}" : $"Error: {result.message}";
    }

    [McpServerTool, Description("Navigate a standard Windows file dialog to select a file. Splits the path into folder + filename, navigates to the folder first, then selects the file. The file dialog must already be open.")]
    public static async Task<string> wpf_file_dialog(
        [Description("Full path to the file to select (e.g., 'C:\\Data\\project.xer')")] string file_path)
    {
        if (Proxy != null)
            return FormatResponse(await Proxy.CallAsync(Constants.Commands.FileDialog, new() { ["filePath"] = file_path }));

        var result = UiaEngine.Instance.FileDialogSetPath(file_path);
        return result.success ? $"OK: {result.message}" : $"Error: {result.message}";
    }

    [McpServerTool, Description("Send keyboard input. Use '+' for simultaneous keys (e.g., 'Ctrl+S'), comma for sequential keys (e.g., 'Alt,F' to activate ribbon keytips then press F).")]
    public static async Task<string> wpf_send_keys(
        [Description("Keys to send (e.g., 'Ctrl+S', 'Alt,F', 'Enter', 'Escape')")] string keys)
    {
        if (Proxy != null)
            return FormatResponse(await Proxy.CallAsync(Constants.Commands.SendKeys, new() { ["keys"] = keys }));

        var result = UiaEngine.Instance.SendKeyboardShortcut(keys);
        return result.success ? $"OK: {result.message}" : $"Error: {result.message}";
    }

    [McpServerTool, Description("Focus the main application window.")]
    public static async Task<string> wpf_focus()
    {
        if (Proxy != null)
            return FormatResponse(await Proxy.CallAsync(Constants.Commands.Focus));

        var result = UiaEngine.Instance.FocusWindow();
        return result.success ? $"OK: {result.message}" : $"Error: {result.message}";
    }

    [McpServerTool, Description("Take a screenshot of the attached application window.")]
    public static async Task<string> wpf_screenshot()
    {
        if (Proxy != null)
        {
            var resp = await Proxy.CallAsync(Constants.Commands.Screenshot);
            if (resp.TryGetProperty("ok", out var ok) && ok.GetBoolean())
            {
                var msg = resp.GetProperty("result").GetString();
                var b64len = resp.TryGetProperty("base64", out var b64) ? b64.GetString()?.Length ?? 0 : 0;
                return $"OK: {msg}\n[Screenshot data: {b64len} chars base64 PNG]";
            }
            return $"Error: {resp.GetProperty("error").GetString()}";
        }

        var result = UiaEngine.Instance.TakeScreenshot();
        if (!result.success) return $"Error: {result.message}";
        return $"OK: {result.message}\n[Screenshot data: {result.base64.Length} chars base64 PNG]";
    }

    [McpServerTool, Description("Get detailed properties of a cached element.")]
    public static async Task<string> wpf_properties(
        [Description("Element reference (e.g., 'e1')")] string ref_key)
    {
        if (Proxy != null)
        {
            var resp = await Proxy.CallAsync(Constants.Commands.Properties, new() { ["refKey"] = ref_key });
            if (resp.TryGetProperty("ok", out var ok) && ok.GetBoolean())
                return JsonSerializer.Serialize(resp.GetProperty("result"), Constants.IndentedJson);
            return $"Error: {resp.GetProperty("error").GetString()}";
        }

        if (!_cache.TryGet(ref_key, out var element))
            return $"Error: Unknown element reference '{ref_key}'";
        var props = UiaEngine.Instance.GetElementProperties(element!);
        return JsonSerializer.Serialize(props, Constants.IndentedJson);
    }

    [McpServerTool, Description("Check if the server is attached to a process. Returns attachment status, window title, and PID.")]
    public static async Task<string> wpf_status()
    {
        if (Proxy != null)
        {
            var resp = await Proxy.CallAsync(Constants.Commands.Status);
            if (resp.TryGetProperty("ok", out var ok) && ok.GetBoolean())
            {
                var attached = resp.GetProperty("attached").GetBoolean();
                if (!attached) return "Not attached to any process. Call wpf_attach first.";
                var title = resp.TryGetProperty("windowTitle", out var wt) && wt.ValueKind == JsonValueKind.String ? wt.GetString() : "unknown";
                var pid = resp.TryGetProperty("pid", out var pp) && pp.ValueKind == JsonValueKind.Number ? pp.GetInt32() : 0;
                return $"Attached to PID {pid}: {title}";
            }
            return $"Error: {resp.GetProperty("error").GetString()}";
        }

        var engine = UiaEngine.Instance;
        if (!engine.IsAttached) return "Not attached to any process. Call wpf_attach first.";
        return $"Attached to PID {engine.ProcessId}: {engine.WindowTitle}";
    }

    /// <summary>Shared macro engine instance. Loaded once from the macros/ folder.</summary>
    private static readonly Lazy<MacroEngine> _macroEngine = new(() => new MacroEngine());

    [McpServerTool, Description("List all available macros with their descriptions and parameters. Also shows knowledge base summaries for supported applications.")]
    public static async Task<string> wpf_macro_list()
    {
        if (Proxy != null)
        {
            var resp = await Proxy.CallAsync(Constants.Commands.MacroList);
            if (resp.TryGetProperty("ok", out var ok) && ok.GetBoolean())
            {
                var output = JsonSerializer.Serialize(resp.GetProperty("result"), Constants.IndentedJson);
                if (resp.TryGetProperty("loadErrors", out var errors) && errors.GetArrayLength() > 0)
                    output += $"\n\nLoad Errors:\n{JsonSerializer.Serialize(errors, Constants.IndentedJson)}";
                if (resp.TryGetProperty("knowledgeBases", out var kbs) && kbs.GetArrayLength() > 0)
                {
                    output += "\n\nKnowledge Bases:";
                    foreach (var kb in kbs.EnumerateArray())
                    {
                        var summary = kb.TryGetProperty("summary", out var s) ? s.GetString() : "";
                        if (!string.IsNullOrEmpty(summary))
                            output += $"\n{summary}";
                    }
                }
                return output;
            }
            return $"Error: {resp.GetProperty("error").GetString()}";
        }

        var engine = _macroEngine.Value;
        var macros = engine.List();
        var loadErrors = engine.LoadErrors;
        var knowledgeBases = engine.KnowledgeBases;

        if (macros.Count == 0 && loadErrors.Count == 0 && knowledgeBases.Count == 0)
            return "No macros found. Place .yaml files in the macros/ folder.";

        var result = macros.Count > 0
            ? JsonSerializer.Serialize(macros, Constants.IndentedJson)
            : "No valid macros loaded.";

        if (loadErrors.Count > 0)
            result += $"\n\nLoad Errors ({loadErrors.Count}):\n{JsonSerializer.Serialize(loadErrors, Constants.IndentedJson)}";

        if (knowledgeBases.Count > 0)
        {
            result += "\n\nKnowledge Bases:";
            foreach (var kb in knowledgeBases)
                result += $"\n{kb.Summary}";
        }

        return result;
    }

    /// <summary>
    /// Provides access to the knowledge base summaries for static resource registration.
    /// </summary>
    internal static IReadOnlyList<KnowledgeBase> GetKnowledgeBases() =>
        _macroEngine.Value.KnowledgeBases;

    [McpServerTool, Description("Run a named macro with optional parameters. Use wpf_macro_list to see available macros.")]
    public static async Task<string> wpf_macro(
        [Description("Macro name (e.g., 'acumen-fuse/import-xer')")] string name,
        [Description("Parameters as JSON object (e.g., '{\"filePath\":\"C:\\\\data\\\\test.xer\"}')")] string? parameters = null)
    {
        if (Proxy != null)
        {
            var args = new Dictionary<string, object?> { ["name"] = name, ["parameters"] = parameters };
            var resp = await Proxy.CallAsync(Constants.Commands.Macro, args);
            if (resp.TryGetProperty("ok", out var ok) && ok.GetBoolean())
                return JsonSerializer.Serialize(resp.GetProperty("result"), Constants.IndentedJson);
            return $"Error: {resp.GetProperty("error").GetString()}";
        }

        var parsedParams = new Dictionary<string, string>();
        if (!string.IsNullOrEmpty(parameters))
        {
            try
            {
                var jsonParams = JsonSerializer.Deserialize<Dictionary<string, string>>(parameters);
                if (jsonParams != null) parsedParams = jsonParams;
            }
            catch (Exception ex)
            {
                return $"Error: Invalid parameters JSON: {ex.Message}";
            }
        }

        var engine = _macroEngine.Value;
        var result = await engine.ExecuteAsync(name, parsedParams);
        return JsonSerializer.Serialize(result, Constants.IndentedJson);
    }

    [McpServerTool, Description("Save a macro YAML file from a list of steps. Auto-derives the product folder from the attached process. Use this to persist workflows as reusable macros.")]
    public static async Task<string> wpf_save_macro(
        [Description("Macro name without product prefix (e.g., 'import-xer'). Product folder is auto-derived from the attached process.")] string name,
        [Description("Human-readable description of what the macro does")] string description,
        [Description("JSON array of step objects. Each step must have an 'action' field. Example: [{\"action\":\"find\",\"automation_id\":\"myBtn\",\"save_as\":\"btn\"},{\"action\":\"click\",\"ref\":\"btn\"}]")] string steps,
        [Description("JSON array of parameter objects with name, description, required, default fields (optional)")] string? parameters = null,
        [Description("Macro timeout in seconds (default 30)")] int timeout = 30,
        [Description("Overwrite existing macro file if it exists (default false)")] bool force = false)
    {
        if (Proxy != null)
        {
            var args = new Dictionary<string, object?>
            {
                ["name"] = name,
                ["description"] = description,
                ["steps"] = steps,
                ["parameters"] = parameters,
                ["timeout"] = timeout,
                ["force"] = force
            };
            var resp = await Proxy.CallAsync(Constants.Commands.SaveMacro, args);
            if (resp.TryGetProperty("ok", out var ok) && ok.GetBoolean())
            {
                var filePath = resp.TryGetProperty("filePath", out var fp) ? fp.GetString() : "";
                var macroName = resp.TryGetProperty("macroName", out var mn) ? mn.GetString() : "";
                var message = resp.TryGetProperty("message", out var msg) ? msg.GetString() : "";
                return $"OK: {message}\nMacro: {macroName}\nFile: {filePath}";
            }
            return $"Error: {resp.GetProperty("error").GetString()}";
        }

        // Direct mode: get attached process name for product folder derivation
        var engine = UiaEngine.Instance;
        if (!engine.IsAttached)
            return "Error: Not attached to any process. Call wpf_attach first.";

        var processName = engine.ProcessName;

        // Parse steps JSON
        List<Dictionary<string, object>> parsedSteps;
        try
        {
            parsedSteps = JsonHelpers.ParseJsonArray(steps);
        }
        catch (Exception ex)
        {
            return $"Error: Invalid steps JSON: {ex.Message}";
        }

        // Parse parameters JSON
        List<Dictionary<string, object>>? parsedParams = null;
        if (!string.IsNullOrEmpty(parameters))
        {
            try
            {
                parsedParams = JsonHelpers.ParseJsonArray(parameters); // same format: array of objects
            }
            catch (Exception ex)
            {
                return $"Error: Invalid parameters JSON: {ex.Message}";
            }
        }

        var macroEngine = _macroEngine.Value;
        var result = macroEngine.SaveMacro(name, description, parsedSteps, parsedParams, timeout, force, processName);

        if (result.Ok)
            return $"OK: {result.Message}\nMacro: {result.MacroName}\nFile: {result.FilePath}";
        return $"Error: {result.Message}";
    }

    [McpServerTool, Description("Export a macro as a Windows shortcut (.lnk) file that can be double-clicked to run. The shortcut runs with administrator privileges. Use 'all' as the name to export all macros.")]
    public static async Task<string> wpf_export_macro(
        [Description("Macro name (e.g., 'acumen-fuse/import-xer') or 'all' to export all macros")] string name,
        [Description("Output directory for shortcuts (optional, defaults to Shortcuts/ next to exe)")] string? output_path = null,
        [Description("Overwrite existing shortcut files (default false)")] bool force = false)
    {
        if (Proxy != null)
        {
            if (name.Equals("all", StringComparison.OrdinalIgnoreCase))
            {
                var allArgs = new Dictionary<string, object?>
                {
                    ["outputPath"] = output_path,
                    ["force"] = force
                };
                var allResp = await Proxy.CallAsync(Constants.Commands.ExportAllMacros, allArgs);
                if (allResp.TryGetProperty("ok", out var allOk) && allOk.GetBoolean())
                {
                    var results = allResp.GetProperty("results");
                    var lines = new List<string>();
                    int success = 0, failed = 0;
                    foreach (var r in results.EnumerateArray())
                    {
                        var rOk = r.GetProperty("ok").GetBoolean();
                        var rName = r.GetProperty("macroName").GetString();
                        var rMsg = r.GetProperty("message").GetString();
                        lines.Add(rOk ? $"  OK: {rName} -> {r.GetProperty("shortcutPath").GetString()}" : $"  FAILED: {rName} - {rMsg}");
                        if (rOk) success++; else failed++;
                    }
                    return $"Exported {success} shortcut(s), {failed} failed:\n{string.Join("\n", lines)}";
                }
                return $"Error: {allResp.GetProperty("error").GetString()}";
            }

            var args = new Dictionary<string, object?>
            {
                ["name"] = name,
                ["outputPath"] = output_path,
                ["force"] = force
            };
            var resp = await Proxy.CallAsync(Constants.Commands.ExportMacro, args);
            if (resp.TryGetProperty("ok", out var ok) && ok.GetBoolean())
            {
                var shortcutPath = resp.TryGetProperty("shortcutPath", out var sp) ? sp.GetString() : "";
                var macroName = resp.TryGetProperty("macroName", out var mn) ? mn.GetString() : "";
                var message = resp.TryGetProperty("message", out var msg) ? msg.GetString() : "";
                return $"OK: {message}\nMacro: {macroName}\nShortcut: {shortcutPath}";
            }
            return $"Error: {resp.GetProperty("error").GetString()}";
        }

        // Direct mode
        var macroEngine = _macroEngine.Value;

        if (name.Equals("all", StringComparison.OrdinalIgnoreCase))
        {
            var results = macroEngine.ExportAllMacros(output_path, force);
            var lines = new List<string>();
            int successCount = 0, failedCount = 0;
            foreach (var r in results)
            {
                lines.Add(r.Ok ? $"  OK: {r.MacroName} -> {r.ShortcutPath}" : $"  FAILED: {r.MacroName} - {r.Message}");
                if (r.Ok) successCount++; else failedCount++;
            }
            return $"Exported {successCount} shortcut(s), {failedCount} failed:\n{string.Join("\n", lines)}";
        }

        var result = macroEngine.ExportMacro(name, output_path, force);
        if (result.Ok)
            return $"OK: {result.Message}\nMacro: {result.MacroName}\nShortcut: {result.ShortcutPath}";
        return $"Error: {result.Message}";
    }

    [McpServerTool, Description("Start a watch session that records focus changes, hover changes, and property changes in the attached WPF application. Use this to observe what a user is doing, then replay or optimize their workflow. Only one session can be active at a time.")]
    public static async Task<string> wpf_watch_start()
    {
        if (Proxy != null)
        {
            var resp = await Proxy.CallAsync(Constants.Commands.WatchStart, new Dictionary<string, object?>());
            if (resp.TryGetProperty("ok", out var ok) && ok.GetBoolean())
            {
                var sessionId = resp.TryGetProperty("sessionId", out var sid) ? sid.GetString() : "";
                return $"OK: Watch started (session {sessionId}). Focus, hover, and property changes are now being recorded.\nCall wpf_watch_stop to end the session and retrieve the recorded entries.";
            }
            return $"Error: {resp.GetProperty("error").GetString()}";
        }
        return "Error: Watch mode requires the elevated server (proxy mode)";
    }

    [McpServerTool, Description("Stop the current watch session and return all recorded entries. Each entry includes timestamps, element details (ControlType, AutomationId, Name), property values, and change diffs. Use this data to understand user workflows, create macros, or optimize interactions.")]
    public static async Task<string> wpf_watch_stop()
    {
        if (Proxy != null)
        {
            var resp = await Proxy.CallAsync(Constants.Commands.WatchStop, new Dictionary<string, object?>());
            if (resp.TryGetProperty("ok", out var ok) && ok.GetBoolean())
                return FormatWatchSession(resp);
            return $"Error: {resp.GetProperty("error").GetString()}";
        }
        return "Error: Watch mode requires the elevated server (proxy mode)";
    }

    [McpServerTool, Description("Get the current or last watch session's entries. Returns the full session log including all focus, hover, and property change events with timestamps and element details.")]
    public static async Task<string> wpf_watch_status()
    {
        if (Proxy != null)
        {
            var resp = await Proxy.CallAsync(Constants.Commands.WatchStatus, new Dictionary<string, object?>());
            if (resp.TryGetProperty("ok", out var ok) && ok.GetBoolean())
                return FormatWatchSession(resp);
            return $"Error: {resp.GetProperty("error").GetString()}";
        }
        return "Error: Watch mode requires the elevated server (proxy mode)";
    }

    private static string FormatWatchSession(JsonElement resp)
    {
        var sessionId = resp.TryGetProperty("sessionId", out var sid) ? sid.GetString() : "?";
        var isActive = resp.TryGetProperty("isActive", out var ia) && ia.GetBoolean();
        var entryCount = resp.TryGetProperty("entryCount", out var ec) ? ec.GetInt32() : 0;

        var startTime = resp.TryGetProperty("startTime", out var st)
            ? DateTime.Parse(st.GetString()!).ToString("HH:mm:ss") : "?";
        var stopTime = resp.TryGetProperty("stopTime", out var et) && et.ValueKind == JsonValueKind.String
            ? DateTime.Parse(et.GetString()!).ToString("HH:mm:ss") : (isActive ? "active" : "?");

        var sb = new System.Text.StringBuilder();
        sb.AppendLine($"Session: {sessionId} | Started: {startTime} | Stopped: {stopTime} | Entries: {entryCount}");
        sb.AppendLine();

        if (resp.TryGetProperty("entries", out var entries))
        {
            foreach (var e in entries.EnumerateArray())
            {
                var time = DateTime.Parse(e.GetProperty("time").GetString()!).ToString("HH:mm:ss.fff");
                var kind = e.GetProperty("kind").GetString();
                var ct = e.TryGetProperty("controlType", out var ctv) ? ctv.GetString() : "";
                var aid = e.TryGetProperty("automationId", out var aidv) ? aidv.GetString() : "";
                var name = e.TryGetProperty("name", out var nv) ? nv.GetString() : "";

                var elementDesc = !string.IsNullOrEmpty(aid) ? $"[{ct}] #{aid}"
                    : !string.IsNullOrEmpty(name) ? $"[{ct}] \"{name}\""
                    : $"[{ct}]";

                if (kind == "PropertyChange")
                {
                    var prop = e.TryGetProperty("changedProperty", out var cpv) ? cpv.GetString() : "?";
                    var oldVal = e.TryGetProperty("oldValue", out var ov) ? ov.GetString() : "";
                    var newVal = e.TryGetProperty("newValue", out var nwv) ? nwv.GetString() : "";
                    sb.AppendLine($"{time} [PropChange] {elementDesc}  {prop}: \"{oldVal}\" → \"{newVal}\"");
                }
                else
                {
                    // Focus or Hover — include key property values
                    var extras = new List<string>();
                    if (e.TryGetProperty("properties", out var props))
                    {
                        if (props.TryGetProperty("Value", out var val) && val.GetString() is string v && v.Length > 0)
                            extras.Add($"Value=\"{(v.Length > 40 ? v[..40] + "..." : v)}\"");
                        if (props.TryGetProperty("ToggleState", out var ts))
                            extras.Add($"Toggle={ts.GetString()}");
                        if (props.TryGetProperty("IsSelected", out var sel) && sel.GetString() == "True")
                            extras.Add("Selected");
                        if (props.TryGetProperty("ExpandCollapseState", out var ecs) && ecs.GetString() != "LeafNode")
                            extras.Add(ecs.GetString()!);
                    }
                    var suffix = extras.Count > 0 ? $"  ({string.Join(", ", extras)})" : "";
                    sb.AppendLine($"{time} [{kind}]  {elementDesc}{suffix}");
                }
            }
        }

        return sb.ToString().TrimEnd();
    }

    // Helper: format simple ok/error response from proxy
    private static string FormatResponse(JsonElement resp)
    {
        if (resp.TryGetProperty("ok", out var ok) && ok.GetBoolean())
            return $"OK: {resp.GetProperty("result").GetString()}";
        return $"Error: {resp.GetProperty("error").GetString()}";
    }

    // Helper: format find response with refKey and properties
    private static string FormatProxyFindResponse(JsonElement resp)
    {
        if (resp.TryGetProperty("ok", out var ok) && ok.GetBoolean())
        {
            var refKey = resp.GetProperty("refKey").GetString();
            var desc = resp.GetProperty("desc").GetString();
            var props = resp.GetProperty("properties");
            return $"OK [{refKey}]: {desc}\nProperties: {JsonSerializer.Serialize(props, Constants.IndentedJson)}";
        }
        return $"Error: {resp.GetProperty("error").GetString()}";
    }
}
