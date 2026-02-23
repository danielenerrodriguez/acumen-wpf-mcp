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

    [McpServerTool, Description("Attach to a running WPF application by process name or PID. Must be called before any other tool.")]
    public static async Task<string> wpf_attach(
        [Description("Process name (without .exe)")] string? process_name = null,
        [Description("Process ID")] int? pid = null)
    {
        if (process_name == null && pid == null)
            return "Error: Either process_name or pid must be provided";

        if (Proxy != null)
        {
            var args = new Dictionary<string, object?> { ["processName"] = process_name, ["pid"] = pid };
            var resp = await Proxy.CallAsync("attach", args);
            return FormatResponse(resp);
        }

        var engine = UiaEngine.Instance;
        var result = pid.HasValue ? engine.AttachByPid(pid.Value) : engine.Attach(process_name!);
        return result.success ? $"OK: {result.message}" : $"Error: {result.message}";
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
