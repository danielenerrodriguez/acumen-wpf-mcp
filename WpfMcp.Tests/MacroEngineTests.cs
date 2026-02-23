using System.IO;
using Xunit;

namespace WpfMcp.Tests;

public class MacroEngineTests : IDisposable
{
    private readonly string _tempDir;

    public MacroEngineTests()
    {
        _tempDir = Path.Combine(Path.GetTempPath(), $"wpfmcp_tests_{Guid.NewGuid():N}");
        Directory.CreateDirectory(_tempDir);
    }

    public void Dispose()
    {
        try { Directory.Delete(_tempDir, recursive: true); } catch { }
    }

    private void WriteMacro(string relativePath, string yaml)
    {
        var fullPath = Path.Combine(_tempDir, relativePath);
        Directory.CreateDirectory(Path.GetDirectoryName(fullPath)!);
        File.WriteAllText(fullPath, yaml);
    }

    // --- Loading & Discovery ---

    [Fact]
    public void Load_ValidYaml_ParsesMacroDefinition()
    {
        WriteMacro("test.yaml", @"
name: Test Macro
description: A test macro
timeout: 30
parameters:
  - name: filePath
    description: Path to file
    required: true
  - name: mode
    required: false
    default: fast
steps:
  - action: focus
  - action: send_keys
    keys: Alt,F
  - action: wait
    seconds: 1.5
");

        using var engine = new MacroEngine(_tempDir, enableWatcher: false);
        var macros = engine.List();
        Assert.Single(macros);
        Assert.Equal("test", macros[0].Name);
        Assert.Equal("Test Macro", macros[0].DisplayName);
        Assert.Equal("A test macro", macros[0].Description);
        Assert.Equal(2, macros[0].Parameters.Count);
        Assert.True(macros[0].Parameters[0].Required);
        Assert.Equal("fast", macros[0].Parameters[1].Default);
    }

    [Fact]
    public void Load_SubfolderMacro_UsesRelativePath()
    {
        WriteMacro(Path.Combine("acumen-fuse", "open-menu.yaml"), @"
name: Open Menu
description: Opens the menu
steps:
  - action: focus
");

        using var engine = new MacroEngine(_tempDir, enableWatcher: false);
        var macros = engine.List();
        Assert.Single(macros);
        Assert.Equal("acumen-fuse/open-menu", macros[0].Name);
    }

    [Fact]
    public void Load_EmptyDirectory_ReturnsNoMacros()
    {
        using var engine = new MacroEngine(_tempDir, enableWatcher: false);
        Assert.Empty(engine.List());
    }

    [Fact]
    public void Load_NonexistentDirectory_ReturnsNoMacros()
    {
        using var engine = new MacroEngine(Path.Combine(_tempDir, "doesnotexist"), enableWatcher: false);
        Assert.Empty(engine.List());
    }

    [Fact]
    public void Load_InvalidYaml_SkipsFile()
    {
        WriteMacro("bad.yaml", "this is not valid yaml: [[[");
        WriteMacro("good.yaml", @"
name: Good
description: Works
steps:
  - action: focus
");

        using var engine = new MacroEngine(_tempDir, enableWatcher: false);
        var macros = engine.List();
        // Bad yaml is skipped, good one is loaded
        Assert.True(macros.Count >= 1);
        Assert.Contains(macros, m => m.Name == "good");
    }

    [Fact]
    public void Get_ExistingMacro_ReturnsDefinition()
    {
        WriteMacro("test.yaml", @"
name: Test
description: Test
steps:
  - action: focus
");

        using var engine = new MacroEngine(_tempDir, enableWatcher: false);
        var macro = engine.Get("test");
        Assert.NotNull(macro);
        Assert.Equal("Test", macro!.Name);
    }

    [Fact]
    public void Get_NonexistentMacro_ReturnsNull()
    {
        using var engine = new MacroEngine(_tempDir, enableWatcher: false);
        Assert.Null(engine.Get("nonexistent"));
    }

    // --- Parameter Validation ---

    [Fact]
    public void ValidateParameters_MissingRequired_ReturnsError()
    {
        var macro = new MacroDefinition
        {
            Parameters = new List<MacroParameter>
            {
                new() { Name = "filePath", Required = true },
                new() { Name = "mode", Required = false, Default = "fast" }
            }
        };

        var error = MacroEngine.ValidateParameters(macro, new Dictionary<string, string>());
        Assert.NotNull(error);
        Assert.Contains("filePath", error);
    }

    [Fact]
    public void ValidateParameters_AllProvided_ReturnsNull()
    {
        var macro = new MacroDefinition
        {
            Parameters = new List<MacroParameter>
            {
                new() { Name = "filePath", Required = true }
            }
        };

        var error = MacroEngine.ValidateParameters(macro, new Dictionary<string, string>
        {
            ["filePath"] = "C:\\test.xer"
        });
        Assert.Null(error);
    }

    [Fact]
    public void ValidateParameters_RequiredWithDefault_NoError()
    {
        var macro = new MacroDefinition
        {
            Parameters = new List<MacroParameter>
            {
                new() { Name = "mode", Required = true, Default = "fast" }
            }
        };

        var error = MacroEngine.ValidateParameters(macro, new Dictionary<string, string>());
        Assert.Null(error);
    }

    // --- Parameter Substitution ---

    [Fact]
    public void SubstituteParams_ReplacesPlaceholders()
    {
        var result = MacroEngine.SubstituteParams(
            "Import {{filePath}} with mode {{mode}}",
            new Dictionary<string, string>
            {
                ["filePath"] = "C:\\test.xer",
                ["mode"] = "fast"
            });
        Assert.Equal("Import C:\\test.xer with mode fast", result);
    }

    [Fact]
    public void SubstituteParams_UnknownParam_LeavesPlaceholder()
    {
        var result = MacroEngine.SubstituteParams(
            "Value is {{unknown}}",
            new Dictionary<string, string>());
        Assert.Equal("Value is {{unknown}}", result);
    }

    [Fact]
    public void SubstituteParams_NullTemplate_ReturnsNull()
    {
        var result = MacroEngine.SubstituteParams(null, new Dictionary<string, string>());
        Assert.Null(result);
    }

    [Fact]
    public void SubstituteParams_NoPlaceholders_ReturnsOriginal()
    {
        var result = MacroEngine.SubstituteParams(
            "No placeholders here",
            new Dictionary<string, string> { ["key"] = "value" });
        Assert.Equal("No placeholders here", result);
    }

    // --- Execution: Non-attached errors ---

    [Fact]
    public async Task Execute_MacroNotFound_ReturnsError()
    {
        using var engine = new MacroEngine(_tempDir, enableWatcher: false);
        var result = await engine.ExecuteAsync("nonexistent");
        Assert.False(result.Success);
        Assert.Contains("not found", result.Message);
    }

    [Fact]
    public async Task Execute_MissingRequiredParam_ReturnsError()
    {
        WriteMacro("test.yaml", @"
name: Test
description: Test
parameters:
  - name: filePath
    required: true
steps:
  - action: focus
");

        using var engine = new MacroEngine(_tempDir, enableWatcher: false);
        var result = await engine.ExecuteAsync("test");
        Assert.False(result.Success);
        Assert.Contains("filePath", result.Message);
    }

    // --- Step parsing ---

    [Fact]
    public void Load_AllStepTypes_ParsesCorrectly()
    {
        WriteMacro("all-steps.yaml", @"
name: All Steps
description: Tests all step types
steps:
  - action: focus
  - action: attach
    process_name: Notepad
  - action: snapshot
    max_depth: 5
  - action: find
    automation_id: myId
    name: myName
    class_name: myClass
    control_type: Button
    save_as: btn
    timeout: 10
    retry_interval: 2
  - action: find_by_path
    path:
      - SearchProp:ControlType~Custom
      - SearchProp:AutomationId~myId
    save_as: pathEl
  - action: click
    ref: btn
  - action: right_click
    ref: btn
  - action: type
    text: Hello {{name}}
  - action: send_keys
    keys: Ctrl+S
  - action: wait
    seconds: 2.5
  - action: screenshot
  - action: properties
    ref: btn
  - action: children
    ref: btn
  - action: macro
    macro_name: other-macro
    params:
      key: value
");

        using var engine = new MacroEngine(_tempDir, enableWatcher: false);
        var macro = engine.Get("all-steps");
        Assert.NotNull(macro);
        Assert.Equal(14, macro!.Steps.Count);
        Assert.Equal("focus", macro.Steps[0].Action);
        Assert.Equal("Notepad", macro.Steps[1].ProcessName);
        Assert.Equal(5, macro.Steps[2].MaxDepth);
        Assert.Equal("myId", macro.Steps[3].AutomationId);
        Assert.Equal("btn", macro.Steps[3].SaveAs);
        Assert.Equal(10, macro.Steps[3].StepTimeout);
        Assert.Equal(2, macro.Steps[3].RetryInterval);
        Assert.Equal(2, macro.Steps[4].Path?.Count);
        Assert.Equal("btn", macro.Steps[5].Ref);
        Assert.Equal("Hello {{name}}", macro.Steps[7].Text);
        Assert.Equal("Ctrl+S", macro.Steps[8].Keys);
        Assert.Equal(2.5, macro.Steps[9].Seconds);
        Assert.Equal("other-macro", macro.Steps[13].MacroName);
        Assert.Equal("value", macro.Steps[13].Params?["key"]);
    }

    [Fact]
    public void Reload_PicksUpNewFiles()
    {
        using var engine = new MacroEngine(_tempDir, enableWatcher: false);
        Assert.Empty(engine.List());

        WriteMacro("new.yaml", @"
name: New
description: Added after load
steps:
  - action: focus
");

        engine.Reload();
        Assert.Single(engine.List());
    }

    // --- Load Errors ---

    [Fact]
    public void LoadErrors_InvalidYaml_PopulatesErrors()
    {
        WriteMacro("bad.yaml", "this is not valid yaml: [[[");

        using var engine = new MacroEngine(_tempDir, enableWatcher: false);
        Assert.Empty(engine.List());
        Assert.Single(engine.LoadErrors);
        Assert.Equal("bad", engine.LoadErrors[0].MacroName);
        Assert.Contains("bad.yaml", engine.LoadErrors[0].FilePath);
        Assert.NotEmpty(engine.LoadErrors[0].Error);
    }

    [Fact]
    public void LoadErrors_EmptyFile_PopulatesErrors()
    {
        WriteMacro("empty.yaml", "");

        using var engine = new MacroEngine(_tempDir, enableWatcher: false);
        Assert.Empty(engine.List());
        Assert.Single(engine.LoadErrors);
        Assert.Contains("null", engine.LoadErrors[0].Error, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void LoadErrors_NoSteps_PopulatesErrors()
    {
        WriteMacro("nosteps.yaml", @"
name: No Steps
description: Missing steps
");

        using var engine = new MacroEngine(_tempDir, enableWatcher: false);
        Assert.Empty(engine.List());
        Assert.Single(engine.LoadErrors);
        Assert.Contains("no steps", engine.LoadErrors[0].Error, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void LoadErrors_MixedFiles_ReportsOnlyBadOnes()
    {
        WriteMacro("good.yaml", @"
name: Good
description: Works
steps:
  - action: focus
");
        WriteMacro("bad1.yaml", "not yaml: [[[");
        WriteMacro("bad2.yaml", "");

        using var engine = new MacroEngine(_tempDir, enableWatcher: false);
        Assert.Single(engine.List());
        Assert.Equal(2, engine.LoadErrors.Count);
    }

    [Fact]
    public void LoadErrors_ClearedOnReload()
    {
        WriteMacro("bad.yaml", "not yaml: [[[");

        using var engine = new MacroEngine(_tempDir, enableWatcher: false);
        Assert.Single(engine.LoadErrors);

        // Fix the file
        WriteMacro("bad.yaml", @"
name: Fixed
description: Now valid
steps:
  - action: focus
");

        engine.Reload();
        Assert.Empty(engine.LoadErrors);
        Assert.Single(engine.List());
    }

    // --- File Watcher ---

    [Fact]
    public async Task Watcher_NewFile_AutoReloads()
    {
        using var engine = new MacroEngine(_tempDir, enableWatcher: true);
        Assert.Empty(engine.List());

        // Add a new macro file
        WriteMacro("watched.yaml", @"
name: Watched
description: Added while watching
steps:
  - action: focus
");

        // Wait for debounce (500ms) + processing time
        await Task.Delay(1500);

        Assert.Single(engine.List());
        Assert.Equal("watched", engine.List()[0].Name);
    }

    [Fact]
    public async Task Watcher_ModifiedFile_AutoReloads()
    {
        WriteMacro("mutable.yaml", @"
name: Original
description: First version
steps:
  - action: focus
");

        using var engine = new MacroEngine(_tempDir, enableWatcher: true);
        Assert.Equal("Original", engine.Get("mutable")!.Name);

        // Modify the file
        WriteMacro("mutable.yaml", @"
name: Updated
description: Second version
steps:
  - action: focus
  - action: wait
    seconds: 1
");

        await Task.Delay(1500);

        Assert.Equal("Updated", engine.Get("mutable")!.Name);
        Assert.Equal(2, engine.Get("mutable")!.Steps.Count);
    }

    [Fact]
    public async Task Watcher_DeletedFile_AutoReloads()
    {
        WriteMacro("doomed.yaml", @"
name: Doomed
description: Will be deleted
steps:
  - action: focus
");

        using var engine = new MacroEngine(_tempDir, enableWatcher: true);
        Assert.Single(engine.List());

        File.Delete(Path.Combine(_tempDir, "doomed.yaml"));

        await Task.Delay(1500);

        Assert.Empty(engine.List());
    }

    [Fact]
    public async Task Watcher_BadFile_ReportsLoadError()
    {
        WriteMacro("good.yaml", @"
name: Good
description: Works
steps:
  - action: focus
");

        using var engine = new MacroEngine(_tempDir, enableWatcher: true);
        Assert.Single(engine.List());
        Assert.Empty(engine.LoadErrors);

        // Add a bad file while watching
        WriteMacro("broken.yaml", "not valid yaml: [[[");

        await Task.Delay(1500);

        // Good macro still loaded, bad one reported as error
        Assert.Single(engine.List());
        Assert.Single(engine.LoadErrors);
        Assert.Equal("broken", engine.LoadErrors[0].MacroName);
    }
}
