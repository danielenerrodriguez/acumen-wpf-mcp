using System.IO;
using YamlDotNet.Serialization;
using YamlDotNet.Serialization.NamingConventions;
using Xunit;

namespace WpfMcp.Tests;

public class MacroSerializerTests : IDisposable
{
    private readonly string _tempDir;

    private static readonly IDeserializer s_yaml = new DeserializerBuilder()
        .WithNamingConvention(UnderscoredNamingConvention.Instance)
        .IgnoreUnmatchedProperties()
        .Build();

    public MacroSerializerTests()
    {
        _tempDir = Path.Combine(Path.GetTempPath(), $"wpfmcp_serializer_{Guid.NewGuid():N}");
        Directory.CreateDirectory(_tempDir);
    }

    public void Dispose()
    {
        try { Directory.Delete(_tempDir, recursive: true); } catch { }
    }

    // --- ToYaml ---

    [Fact]
    public void ToYaml_ValidMacro_ProducesValidYaml()
    {
        var macro = new MacroDefinition
        {
            Name = "Test Macro",
            Description = "A test",
            Steps = new List<MacroStep>
            {
                new() { Action = "focus" },
                new() { Action = "send_keys", Keys = "Ctrl+S" },
            }
        };

        var yaml = MacroSerializer.ToYaml(macro);

        Assert.NotNull(yaml);
        Assert.Contains("name: Test Macro", yaml);
        Assert.Contains("description: A test", yaml);
        Assert.Contains("action: focus", yaml);
        Assert.Contains("keys: Ctrl+S", yaml);
    }

    [Fact]
    public void ToYaml_RoundTrip_DeserializesBackCorrectly()
    {
        var original = new MacroDefinition
        {
            Name = "Round Trip",
            Description = "Testing round trip",
            Steps = new List<MacroStep>
            {
                new() { Action = "find", AutomationId = "uxBtn", ControlType = "Button", SaveAs = "btn1" },
                new() { Action = "click", Ref = "btn1" },
                new() { Action = "type", Text = "Hello World" },
                new() { Action = "send_keys", Keys = "Alt,F" },
                new() { Action = "wait", Seconds = 2.5 },
            }
        };

        var yaml = MacroSerializer.ToYaml(original);
        var deserialized = s_yaml.Deserialize<MacroDefinition>(yaml);

        Assert.Equal(original.Name, deserialized.Name);
        Assert.Equal(original.Description, deserialized.Description);
        Assert.Equal(original.Steps.Count, deserialized.Steps.Count);

        Assert.Equal("find", deserialized.Steps[0].Action);
        Assert.Equal("uxBtn", deserialized.Steps[0].AutomationId);
        Assert.Equal("Button", deserialized.Steps[0].ControlType);
        Assert.Equal("btn1", deserialized.Steps[0].SaveAs);

        Assert.Equal("click", deserialized.Steps[1].Action);
        Assert.Equal("btn1", deserialized.Steps[1].Ref);

        Assert.Equal("type", deserialized.Steps[2].Action);
        Assert.Equal("Hello World", deserialized.Steps[2].Text);

        Assert.Equal("send_keys", deserialized.Steps[3].Action);
        Assert.Equal("Alt,F", deserialized.Steps[3].Keys);

        Assert.Equal("wait", deserialized.Steps[4].Action);
        Assert.Equal(2.5, deserialized.Steps[4].Seconds);
    }

    [Fact]
    public void ToYaml_OmitsDefaultValues()
    {
        var macro = new MacroDefinition
        {
            Name = "Minimal",
            Description = "Test defaults",
            Steps = new List<MacroStep>
            {
                new() { Action = "focus" },
            }
        };

        var yaml = MacroSerializer.ToYaml(macro);

        // Should not contain fields that are null/default
        Assert.DoesNotContain("automation_id:", yaml);
        Assert.DoesNotContain("class_name:", yaml);
        Assert.DoesNotContain("control_type:", yaml);
        Assert.DoesNotContain("ref:", yaml);
        Assert.DoesNotContain("text:", yaml);
        Assert.DoesNotContain("keys:", yaml);
        Assert.DoesNotContain("timeout:", yaml);
    }

    // --- SaveToFile ---

    [Fact]
    public void SaveToFile_CreatesFileAtCorrectPath()
    {
        var macro = new MacroDefinition
        {
            Name = "Test",
            Description = "Test",
            Steps = new List<MacroStep> { new() { Action = "focus" } }
        };

        var filePath = MacroSerializer.SaveToFile(macro, "simple-test", _tempDir);

        Assert.True(File.Exists(filePath));
        Assert.Equal(Path.Combine(_tempDir, "simple-test.yaml"), filePath);

        var content = File.ReadAllText(filePath);
        Assert.Contains("name: Test", content);
    }

    [Fact]
    public void SaveToFile_CreatesSubdirectories()
    {
        var macro = new MacroDefinition
        {
            Name = "Nested",
            Description = "In a subfolder",
            Steps = new List<MacroStep> { new() { Action = "focus" } }
        };

        var filePath = MacroSerializer.SaveToFile(macro, "acumen-fuse/my-workflow", _tempDir);

        Assert.True(File.Exists(filePath));
        Assert.Contains("acumen-fuse", filePath);
        Assert.EndsWith(".yaml", filePath);
        Assert.True(Directory.Exists(Path.Combine(_tempDir, "acumen-fuse")));
    }

    [Fact]
    public void SaveToFile_RoundTrip_LoadableByMacroEngine()
    {
        var macro = new MacroDefinition
        {
            Name = "Loadable",
            Description = "Should be loadable by MacroEngine",
            Steps = new List<MacroStep>
            {
                new() { Action = "find", AutomationId = "btn1", SaveAs = "el1" },
                new() { Action = "click", Ref = "el1" },
            }
        };

        MacroSerializer.SaveToFile(macro, "loadable", _tempDir);

        using var engine = new MacroEngine(_tempDir, enableWatcher: false);
        var loaded = engine.Get("loadable");
        Assert.NotNull(loaded);
        Assert.Equal("Loadable", loaded!.Name);
        Assert.Equal(2, loaded.Steps.Count);
        Assert.Equal("btn1", loaded.Steps[0].AutomationId);
    }

    // --- BuildFromRecordedActions ---

    [Fact]
    public void BuildFromRecordedActions_ClickWithAutomationId_ProducesFindAndClick()
    {
        var actions = new List<RecordedAction>
        {
            new()
            {
                Type = RecordedActionType.Click,
                Timestamp = DateTime.UtcNow,
                X = 100, Y = 200,
                AutomationId = "uxSubmitButton",
                ControlType = "Button",
            }
        };

        var macro = MacroSerializer.BuildFromRecordedActions("Test", "Test description", actions);

        Assert.Equal(2, macro.Steps.Count);

        // Step 0: find
        Assert.Equal("find", macro.Steps[0].Action);
        Assert.Equal("uxSubmitButton", macro.Steps[0].AutomationId);
        Assert.Equal("Button", macro.Steps[0].ControlType);
        Assert.NotNull(macro.Steps[0].SaveAs);

        // Step 1: click referencing the found element
        Assert.Equal("click", macro.Steps[1].Action);
        Assert.Equal(macro.Steps[0].SaveAs, macro.Steps[1].Ref);
    }

    [Fact]
    public void BuildFromRecordedActions_RightClick_ProducesRightClickStep()
    {
        var actions = new List<RecordedAction>
        {
            new()
            {
                Type = RecordedActionType.RightClick,
                Timestamp = DateTime.UtcNow,
                AutomationId = "uxMenu",
                ControlType = "MenuItem",
            }
        };

        var macro = MacroSerializer.BuildFromRecordedActions("Test", "Test", actions);

        Assert.Equal("right_click", macro.Steps[1].Action);
    }

    [Fact]
    public void BuildFromRecordedActions_ClickWithNameAndControlType_FallsBackToName()
    {
        var actions = new List<RecordedAction>
        {
            new()
            {
                Type = RecordedActionType.Click,
                Timestamp = DateTime.UtcNow,
                AutomationId = null,
                ElementName = "OK",
                ControlType = "Button",
            }
        };

        var macro = MacroSerializer.BuildFromRecordedActions("Test", "Test", actions);

        var findStep = macro.Steps[0];
        Assert.Equal("find", findStep.Action);
        Assert.Null(findStep.AutomationId);
        Assert.Equal("OK", findStep.Name);
        Assert.Equal("Button", findStep.ControlType);
    }

    [Fact]
    public void BuildFromRecordedActions_ClickWithClassName_FallsBackToClassName()
    {
        var actions = new List<RecordedAction>
        {
            new()
            {
                Type = RecordedActionType.Click,
                Timestamp = DateTime.UtcNow,
                AutomationId = null,
                ElementName = null,
                ClassName = "XamRibbon",
                ControlType = "Custom",
            }
        };

        var macro = MacroSerializer.BuildFromRecordedActions("Test", "Test", actions);

        var findStep = macro.Steps[0];
        Assert.Null(findStep.AutomationId);
        Assert.Null(findStep.Name);
        Assert.Equal("XamRibbon", findStep.ClassName);
        Assert.Equal("Custom", findStep.ControlType);
    }

    [Fact]
    public void BuildFromRecordedActions_SendKeys_ProducesSendKeysStep()
    {
        var actions = new List<RecordedAction>
        {
            new()
            {
                Type = RecordedActionType.SendKeys,
                Timestamp = DateTime.UtcNow,
                Keys = "Ctrl+S",
            }
        };

        var macro = MacroSerializer.BuildFromRecordedActions("Test", "Test", actions);

        Assert.Single(macro.Steps);
        Assert.Equal("send_keys", macro.Steps[0].Action);
        Assert.Equal("Ctrl+S", macro.Steps[0].Keys);
    }

    [Fact]
    public void BuildFromRecordedActions_Type_ProducesTypeStep()
    {
        var actions = new List<RecordedAction>
        {
            new()
            {
                Type = RecordedActionType.Type,
                Timestamp = DateTime.UtcNow,
                Text = "Hello World",
            }
        };

        var macro = MacroSerializer.BuildFromRecordedActions("Test", "Test", actions);

        Assert.Single(macro.Steps);
        Assert.Equal("type", macro.Steps[0].Action);
        Assert.Equal("Hello World", macro.Steps[0].Text);
    }

    [Fact]
    public void BuildFromRecordedActions_WaitInsertion_AddsWaitStep()
    {
        var baseTime = DateTime.UtcNow;
        var actions = new List<RecordedAction>
        {
            new()
            {
                Type = RecordedActionType.SendKeys,
                Timestamp = baseTime,
                Keys = "Enter",
            },
            new()
            {
                Type = RecordedActionType.SendKeys,
                Timestamp = baseTime.AddSeconds(3),
                Keys = "Tab",
                WaitBeforeSec = 3.0,
            }
        };

        var macro = MacroSerializer.BuildFromRecordedActions("Test", "Test", actions);

        // Should be: send_keys(Enter), wait(3.0), send_keys(Tab)
        Assert.Equal(3, macro.Steps.Count);
        Assert.Equal("send_keys", macro.Steps[0].Action);
        Assert.Equal("wait", macro.Steps[1].Action);
        Assert.Equal(3.0, macro.Steps[1].Seconds);
        Assert.Equal("send_keys", macro.Steps[2].Action);
    }

    [Fact]
    public void BuildFromRecordedActions_SmallWait_NotInserted()
    {
        var baseTime = DateTime.UtcNow;
        var actions = new List<RecordedAction>
        {
            new()
            {
                Type = RecordedActionType.SendKeys,
                Timestamp = baseTime,
                Keys = "A",
            },
            new()
            {
                Type = RecordedActionType.SendKeys,
                Timestamp = baseTime.AddMilliseconds(200),
                Keys = "B",
                WaitBeforeSec = 0.2, // Below 0.5 threshold
            }
        };

        var macro = MacroSerializer.BuildFromRecordedActions("Test", "Test", actions);

        // No wait step inserted for gaps below 0.5s
        Assert.Equal(2, macro.Steps.Count);
        Assert.Equal("send_keys", macro.Steps[0].Action);
        Assert.Equal("send_keys", macro.Steps[1].Action);
    }

    [Fact]
    public void BuildFromRecordedActions_MixedActions_ProducesCorrectSequence()
    {
        var baseTime = DateTime.UtcNow;
        var actions = new List<RecordedAction>
        {
            new()
            {
                Type = RecordedActionType.Click,
                Timestamp = baseTime,
                AutomationId = "uxInput",
                ControlType = "Edit",
            },
            new()
            {
                Type = RecordedActionType.Type,
                Timestamp = baseTime.AddMilliseconds(500),
                Text = "test data",
            },
            new()
            {
                Type = RecordedActionType.SendKeys,
                Timestamp = baseTime.AddSeconds(1),
                Keys = "Ctrl+S",
            }
        };

        var macro = MacroSerializer.BuildFromRecordedActions("Mixed", "Mixed actions test", actions);

        // find + click + type + send_keys = 4 steps
        Assert.Equal(4, macro.Steps.Count);
        Assert.Equal("find", macro.Steps[0].Action);
        Assert.Equal("click", macro.Steps[1].Action);
        Assert.Equal("type", macro.Steps[2].Action);
        Assert.Equal("test data", macro.Steps[2].Text);
        Assert.Equal("send_keys", macro.Steps[3].Action);
        Assert.Equal("Ctrl+S", macro.Steps[3].Keys);
    }

    [Fact]
    public void BuildFromRecordedActions_EmptyActions_ProducesEmptySteps()
    {
        var macro = MacroSerializer.BuildFromRecordedActions("Empty", "No actions", new List<RecordedAction>());

        Assert.Empty(macro.Steps);
        Assert.Equal("Empty", macro.Name);
        Assert.Equal("No actions", macro.Description);
    }

    [Fact]
    public void BuildFromRecordedActions_AliasesAreSequential()
    {
        var actions = new List<RecordedAction>
        {
            new()
            {
                Type = RecordedActionType.Click,
                Timestamp = DateTime.UtcNow,
                AutomationId = "btn1",
                ControlType = "Button",
            },
            new()
            {
                Type = RecordedActionType.Click,
                Timestamp = DateTime.UtcNow.AddSeconds(1),
                AutomationId = "btn2",
                ControlType = "Button",
            }
        };

        var macro = MacroSerializer.BuildFromRecordedActions("Test", "Test", actions);

        // 4 steps: find(step1) + click(step1) + find(step2) + click(step2)
        Assert.Equal(4, macro.Steps.Count);
        Assert.Equal("step1", macro.Steps[0].SaveAs);
        Assert.Equal("step1", macro.Steps[1].Ref);
        Assert.Equal("step2", macro.Steps[2].SaveAs);
        Assert.Equal("step2", macro.Steps[3].Ref);
    }

    [Fact]
    public void BuildFromRecordedActions_FindStepHasTimeout()
    {
        var actions = new List<RecordedAction>
        {
            new()
            {
                Type = RecordedActionType.Click,
                Timestamp = DateTime.UtcNow,
                AutomationId = "uxBtn",
                ControlType = "Button",
            }
        };

        var macro = MacroSerializer.BuildFromRecordedActions("Test", "Test", actions);

        Assert.Equal(Constants.DefaultRecordingStepTimeoutSec, macro.Steps[0].StepTimeout);
    }
}
