using System.IO;
using System.Text;
using YamlDotNet.Serialization;
using YamlDotNet.Serialization.NamingConventions;

namespace WpfMcp;

/// <summary>
/// Serializes MacroDefinition objects to YAML and writes them to disk.
/// Used by the recorder to save captured macros.
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

    /// <summary>
    /// Build a MacroDefinition from a list of recorded actions.
    /// Handles element identification, find+click pairing, typing coalescing,
    /// and wait insertion.
    /// </summary>
    public static MacroDefinition BuildFromRecordedActions(
        string displayName, string description,
        List<RecordedAction> actions)
    {
        var macro = new MacroDefinition
        {
            Name = displayName,
            Description = description,
            Timeout = 0, // use default
        };

        int aliasCounter = 0;

        for (int i = 0; i < actions.Count; i++)
        {
            var action = actions[i];

            // Insert wait step if there's a significant gap before this action
            if (action.WaitBeforeSec.HasValue && action.WaitBeforeSec.Value >= 0.5)
            {
                macro.Steps.Add(new MacroStep
                {
                    Action = "wait",
                    Seconds = Math.Round(action.WaitBeforeSec.Value * 2, MidpointRounding.ToEven) / 2 // round to 0.5s
                });
            }

            switch (action.Type)
            {
                case RecordedActionType.Click:
                case RecordedActionType.RightClick:
                {
                    var alias = $"step{++aliasCounter}";
                    var findStep = BuildFindStep(action, alias);
                    macro.Steps.Add(findStep);
                    macro.Steps.Add(new MacroStep
                    {
                        Action = action.Type == RecordedActionType.Click ? "click" : "right_click",
                        Ref = alias
                    });
                    break;
                }

                case RecordedActionType.SendKeys:
                    macro.Steps.Add(new MacroStep
                    {
                        Action = "send_keys",
                        Keys = action.Keys
                    });
                    break;

                case RecordedActionType.Type:
                    macro.Steps.Add(new MacroStep
                    {
                        Action = "type",
                        Text = action.Text
                    });
                    break;
            }
        }

        return macro;
    }

    /// <summary>Build a find step using the best available element identifier.</summary>
    private static MacroStep BuildFindStep(RecordedAction action, string saveAs)
    {
        var step = new MacroStep
        {
            Action = "find",
            SaveAs = saveAs,
            StepTimeout = Constants.DefaultRecordingStepTimeoutSec,
        };

        // Priority: AutomationId > Name+ControlType > ClassName+ControlType > ControlType alone
        if (!string.IsNullOrEmpty(action.AutomationId))
        {
            step.AutomationId = action.AutomationId;
            // Add ControlType as secondary filter for specificity
            if (!string.IsNullOrEmpty(action.ControlType))
                step.ControlType = action.ControlType;
        }
        else if (!string.IsNullOrEmpty(action.ElementName) && !string.IsNullOrEmpty(action.ControlType))
        {
            step.Name = action.ElementName;
            step.ControlType = action.ControlType;
        }
        else if (!string.IsNullOrEmpty(action.ClassName) && !string.IsNullOrEmpty(action.ControlType))
        {
            step.ClassName = action.ClassName;
            step.ControlType = action.ControlType;
        }
        else if (!string.IsNullOrEmpty(action.ControlType))
        {
            step.ControlType = action.ControlType;
            // Less specific — add any available property
            if (!string.IsNullOrEmpty(action.ElementName))
                step.Name = action.ElementName;
            else if (!string.IsNullOrEmpty(action.ClassName))
                step.ClassName = action.ClassName;
        }
        else
        {
            // Fallback: use whatever is available
            if (!string.IsNullOrEmpty(action.ElementName))
                step.Name = action.ElementName;
            if (!string.IsNullOrEmpty(action.ClassName))
                step.ClassName = action.ClassName;
        }

        return step;
    }
}

/// <summary>Types of actions that can be recorded.</summary>
public enum RecordedActionType
{
    Click,
    RightClick,
    SendKeys,
    Type,
}

/// <summary>A single captured user action during recording.</summary>
public record RecordedAction
{
    public required RecordedActionType Type { get; init; }
    public DateTime Timestamp { get; init; }

    /// <summary>Wait to insert before this action (computed during coalescing).</summary>
    public double? WaitBeforeSec { get; set; }

    // For Click / RightClick
    public int? X { get; init; }
    public int? Y { get; init; }
    public string? AutomationId { get; init; }
    public string? ElementName { get; init; }
    public string? ClassName { get; init; }
    public string? ControlType { get; init; }

    // For SendKeys
    public string? Keys { get; init; }

    // For Type
    public string? Text { get; init; }
}
