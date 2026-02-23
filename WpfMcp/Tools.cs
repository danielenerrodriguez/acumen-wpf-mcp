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
                var launchResp = await Proxy.CallAsync("launch", launchArgs);
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
                    var waitResp = await Proxy.CallAsync("waitForWindow", waitArgs);
                    if (waitResp.TryGetProperty("ok", out var wok) && !wok.GetBoolean())
                        return FormatResponse(waitResp);
                }

                return FormatResponse(launchResp);
            }

            var args = new Dictionary<string, object?> { ["processName"] = process_name, ["pid"] = pid };
            var resp = await Proxy.CallAsync("attach", args);

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
            var resp = await Proxy.CallAsync("snapshot", args);
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
            var resp = await Proxy.CallAsync("find", args);
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
            var resp = await Proxy.CallAsync("findByPath", args);
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
            var resp = await Proxy.CallAsync("children", args);
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
            return FormatResponse(await Proxy.CallAsync("click", new() { ["refKey"] = ref_key }));

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
            return FormatResponse(await Proxy.CallAsync("rightClick", new() { ["refKey"] = ref_key }));

        if (!_cache.TryGet(ref_key, out var element))
            return $"Error: Unknown element reference '{ref_key}'";
        var result = UiaEngine.Instance.RightClickElement(element!);
        return result.success ? $"OK: {result.message}" : $"Error: {result.message}";
    }

    [McpServerTool, Description("Type text into the currently focused element.")]
    public static async Task<string> wpf_type(
        [Description("Text to type")] string text)
    {
        if (Proxy != null)
            return FormatResponse(await Proxy.CallAsync("type", new() { ["text"] = text }));

        var result = UiaEngine.Instance.TypeText(text);
        return result.success ? $"OK: {result.message}" : $"Error: {result.message}";
    }

    [McpServerTool, Description("Send keyboard input. Use '+' for simultaneous keys (e.g., 'Ctrl+S'), comma for sequential keys (e.g., 'Alt,F' to activate ribbon keytips then press F).")]
    public static async Task<string> wpf_send_keys(
        [Description("Keys to send (e.g., 'Ctrl+S', 'Alt,F', 'Enter', 'Escape')")] string keys)
    {
        if (Proxy != null)
            return FormatResponse(await Proxy.CallAsync("sendKeys", new() { ["keys"] = keys }));

        var result = UiaEngine.Instance.SendKeyboardShortcut(keys);
        return result.success ? $"OK: {result.message}" : $"Error: {result.message}";
    }

    [McpServerTool, Description("Focus the main application window.")]
    public static async Task<string> wpf_focus()
    {
        if (Proxy != null)
            return FormatResponse(await Proxy.CallAsync("focus"));

        var result = UiaEngine.Instance.FocusWindow();
        return result.success ? $"OK: {result.message}" : $"Error: {result.message}";
    }

    [McpServerTool, Description("Take a screenshot of the attached application window.")]
    public static async Task<string> wpf_screenshot()
    {
        if (Proxy != null)
        {
            var resp = await Proxy.CallAsync("screenshot");
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
            var resp = await Proxy.CallAsync("properties", new() { ["refKey"] = ref_key });
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
            var resp = await Proxy.CallAsync("status");
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

    [McpServerTool, Description("List all available macros with their descriptions and parameters.")]
    public static async Task<string> wpf_macro_list()
    {
        if (Proxy != null)
        {
            var resp = await Proxy.CallAsync("macroList");
            if (resp.TryGetProperty("ok", out var ok) && ok.GetBoolean())
            {
                var output = JsonSerializer.Serialize(resp.GetProperty("result"), Constants.IndentedJson);
                if (resp.TryGetProperty("loadErrors", out var errors) && errors.GetArrayLength() > 0)
                    output += $"\n\nLoad Errors:\n{JsonSerializer.Serialize(errors, Constants.IndentedJson)}";
                return output;
            }
            return $"Error: {resp.GetProperty("error").GetString()}";
        }

        var engine = _macroEngine.Value;
        var macros = engine.List();
        var loadErrors = engine.LoadErrors;

        if (macros.Count == 0 && loadErrors.Count == 0)
            return "No macros found. Place .yaml files in the macros/ folder.";

        var result = macros.Count > 0
            ? JsonSerializer.Serialize(macros, Constants.IndentedJson)
            : "No valid macros loaded.";

        if (loadErrors.Count > 0)
            result += $"\n\nLoad Errors ({loadErrors.Count}):\n{JsonSerializer.Serialize(loadErrors, Constants.IndentedJson)}";

        return result;
    }

    [McpServerTool, Description("Run a named macro with optional parameters. Use wpf_macro_list to see available macros.")]
    public static async Task<string> wpf_macro(
        [Description("Macro name (e.g., 'acumen-fuse/import-xer')")] string name,
        [Description("Parameters as JSON object (e.g., '{\"filePath\":\"C:\\\\data\\\\test.xer\"}')")] string? parameters = null)
    {
        if (Proxy != null)
        {
            var args = new Dictionary<string, object?> { ["name"] = name, ["parameters"] = parameters };
            var resp = await Proxy.CallAsync("macro", args);
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

    [McpServerTool, Description("Start recording a macro. Captures mouse clicks and keyboard input on the attached application. Call wpf_record_stop to finish.")]
    public static async Task<string> wpf_record_start(
        [Description("Macro name including subfolder (e.g., 'acumen-fuse/my-workflow')")] string name,
        [Description("Path to macros folder (optional, uses WPFMCP_MACROS_PATH or default)")] string? macros_path = null)
    {
        if (Proxy != null)
        {
            var args = new Dictionary<string, object?> { ["name"] = name, ["macrosPath"] = macros_path };
            return FormatResponse(await Proxy.CallAsync("startRecording", args));
        }

        return "Error: Recording requires the elevated server (use --mcp-connect mode)";
    }

    [McpServerTool, Description("Stop recording and save the macro as a YAML file. Returns the generated YAML for review.")]
    public static async Task<string> wpf_record_stop()
    {
        if (Proxy != null)
        {
            var resp = await Proxy.CallAsync("stopRecording");
            if (resp.TryGetProperty("ok", out var ok) && ok.GetBoolean())
            {
                var msg = resp.GetProperty("result").GetString();
                var yaml = resp.TryGetProperty("yaml", out var y) && y.ValueKind == JsonValueKind.String ? y.GetString() : null;
                var filePath = resp.TryGetProperty("filePath", out var fp) && fp.ValueKind == JsonValueKind.String ? fp.GetString() : null;
                var output = $"OK: {msg}";
                if (filePath != null)
                    output += $"\nFile: {filePath}";
                if (yaml != null)
                    output += $"\n\n--- Generated YAML ---\n{yaml}";
                return output;
            }
            return $"Error: {resp.GetProperty("error").GetString()}";
        }

        return "Error: Recording requires the elevated server (use --mcp-connect mode)";
    }

    [McpServerTool, Description("Check the current recording status (idle or recording, action count, duration).")]
    public static async Task<string> wpf_record_status()
    {
        if (Proxy != null)
        {
            var resp = await Proxy.CallAsync("recordingStatus");
            if (resp.TryGetProperty("ok", out var ok) && ok.GetBoolean())
            {
                var state = resp.GetProperty("state").GetString();
                var macroName = resp.TryGetProperty("macroName", out var mn) ? mn.GetString() : "";
                var actionCount = resp.TryGetProperty("actionCount", out var ac) ? ac.GetInt32() : 0;
                var durationSec = resp.TryGetProperty("durationSec", out var ds) ? ds.GetDouble() : 0;

                if (state == "Idle")
                    return "Not recording.";
                return $"Recording: {macroName} ({actionCount} actions, {durationSec:F1}s)";
            }
            return $"Error: {resp.GetProperty("error").GetString()}";
        }

        return "Error: Recording requires the elevated server (use --mcp-connect mode)";
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
