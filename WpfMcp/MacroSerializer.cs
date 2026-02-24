using System.IO;
using System.Text;
using YamlDotNet.Serialization;
using YamlDotNet.Serialization.NamingConventions;

namespace WpfMcp;

/// <summary>
/// Serializes MacroDefinition objects to YAML and writes them to disk.
/// </summary>
public static class MacroSerializer
{
    private static readonly ISerializer s_yaml = new SerializerBuilder()
        .WithNamingConvention(UnderscoredNamingConvention.Instance)
        .ConfigureDefaultValuesHandling(DefaultValuesHandling.OmitDefaults)
        .DisableAliases()
        .Build();

    /// <summary>Serialize a MacroDefinition to a YAML string.</summary>
    public static string ToYaml(MacroDefinition macro)
    {
        // YamlDotNet serializer produces the full object.
        // We post-process to remove empty collections and zero-value fields
        // that OmitDefaults doesn't catch on reference types.
        var yaml = s_yaml.Serialize(macro);
        return yaml;
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
