using System.IO;
using System.Text;

namespace WpfMcp;

/// <summary>
/// Serializes MacroDefinition objects to YAML and writes them to disk.
/// </summary>
public static class MacroSerializer
{
    /// <summary>Serialize a MacroDefinition to a YAML string.</summary>
    public static string ToYaml(MacroDefinition macro)
    {
        // Uses shared serializer from YamlHelpers (underscore naming, omit defaults, no aliases).
        return YamlHelpers.Serializer.Serialize(macro);
    }

    /// <summary>
    /// Save a MacroDefinition to a YAML file.
    /// Creates subdirectories as needed (e.g., "acumen-fuse/my-workflow" → macros/acumen-fuse/my-workflow.yaml).
    /// Returns the full path of the saved file.
    /// </summary>
    public static string SaveToFile(MacroDefinition macro, string macroName, string macrosBasePath)
    {
        var relativePath = macroName.Replace('/', Path.DirectorySeparatorChar) + ".yaml";
        var fullPath = Path.Combine(macrosBasePath, relativePath);

        var dir = Path.GetDirectoryName(fullPath);
        if (!string.IsNullOrEmpty(dir) && !Directory.Exists(dir))
            Directory.CreateDirectory(dir);

        var yaml = ToYaml(macro);
        File.WriteAllText(fullPath, yaml, Encoding.UTF8);
        return fullPath;
    }

}
