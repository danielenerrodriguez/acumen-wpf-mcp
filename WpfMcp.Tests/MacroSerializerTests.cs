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
}
