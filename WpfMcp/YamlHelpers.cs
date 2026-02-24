using YamlDotNet.Serialization;
using YamlDotNet.Serialization.NamingConventions;

namespace WpfMcp;

/// <summary>
/// Shared YAML serialization instances.
/// Centralizes deserializer/serializer configuration so that all callers
/// (Program.cs, CliMode.cs, UiaProxy.cs, MacroEngine.cs) use identical settings.
/// </summary>
public static class YamlHelpers
{
    /// <summary>
    /// Shared deserializer for <see cref="MacroDefinition"/> YAML files.
    /// Uses underscore naming convention and ignores unmatched properties.
    /// </summary>
    public static readonly IDeserializer Deserializer = new DeserializerBuilder()
        .WithNamingConvention(UnderscoredNamingConvention.Instance)
        .IgnoreUnmatchedProperties()
        .Build();

    /// <summary>
    /// Shared deserializer for generic dictionary YAML files (knowledge bases).
    /// Does not use naming conventions — keys are preserved as-is.
    /// </summary>
    public static readonly IDeserializer DictDeserializer = new DeserializerBuilder()
        .IgnoreUnmatchedProperties()
        .Build();

    /// <summary>
    /// Shared serializer for writing macro YAML files.
    /// Uses underscore naming, omits defaults, disables anchors/aliases.
    /// </summary>
    public static readonly ISerializer Serializer = new SerializerBuilder()
        .WithNamingConvention(UnderscoredNamingConvention.Instance)
        .ConfigureDefaultValuesHandling(DefaultValuesHandling.OmitDefaults)
        .DisableAliases()
        .Build();
}
