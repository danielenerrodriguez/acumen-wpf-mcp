using System.Text.Json;

namespace WpfMcp;

/// <summary>
/// Shared JSON utility methods for parsing step/parameter arrays
/// from MCP tool calls and proxy commands.
/// </summary>
public static class JsonHelpers
{
    /// <summary>Parse a JSON array of objects into a list of string-keyed dictionaries.</summary>
    public static List<Dictionary<string, object>> ParseJsonArray(string json)
    {
        var result = new List<Dictionary<string, object>>();
        var doc = JsonDocument.Parse(json);
        foreach (var element in doc.RootElement.EnumerateArray())
        {
            var dict = new Dictionary<string, object>();
            foreach (var prop in element.EnumerateObject())
            {
                dict[prop.Name] = ConvertJsonElement(prop.Value);
            }
            result.Add(dict);
        }
        return result;
    }

    /// <summary>Convert a JsonElement to a plain .NET object for YAML serialization.</summary>
    public static object ConvertJsonElement(JsonElement element)
    {
        return element.ValueKind switch
        {
            JsonValueKind.String => element.GetString()!,
            JsonValueKind.Number => element.TryGetInt32(out var i) ? i : element.GetDouble(),
            JsonValueKind.True => true,
            JsonValueKind.False => false,
            JsonValueKind.Array => element.EnumerateArray().Select(ConvertJsonElement).ToList(),
            JsonValueKind.Object => element.EnumerateObject()
                .ToDictionary(p => p.Name, p => ConvertJsonElement(p.Value)),
            _ => element.ToString()
        };
    }
}
