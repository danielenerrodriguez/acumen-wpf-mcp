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
  - action: launch
    exe_path: ""C:\\MyApp\\App.exe""
    arguments: --verbose
    working_directory: ""C:\\MyApp""
    if_not_running: true
    timeout: 30
  - action: wait_for_window
    title_contains: My App
    timeout: 20
    retry_interval: 2
  - action: verify
    ref: btn
    property: value
    expected: Done
    message: Must be Done to continue
");

        using var engine = new MacroEngine(_tempDir, enableWatcher: false);
        var macro = engine.Get("all-steps");
        Assert.NotNull(macro);
        Assert.Equal(17, macro!.Steps.Count);
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
        // Launch step
        Assert.Equal("launch", macro.Steps[14].Action);
        Assert.Equal("C:\\MyApp\\App.exe", macro.Steps[14].ExePath);
        Assert.Equal("--verbose", macro.Steps[14].Arguments);
        Assert.Equal("C:\\MyApp", macro.Steps[14].WorkingDirectory);
        Assert.True(macro.Steps[14].IfNotRunning);
        Assert.Equal(30, macro.Steps[14].StepTimeout);
        // WaitForWindow step
        Assert.Equal("wait_for_window", macro.Steps[15].Action);
        Assert.Equal("My App", macro.Steps[15].TitleContains);
        Assert.Equal(20, macro.Steps[15].StepTimeout);
        Assert.Equal(2, macro.Steps[15].RetryInterval);
        // Verify step
        Assert.Equal("verify", macro.Steps[16].Action);
        Assert.Equal("btn", macro.Steps[16].Ref);
        Assert.Equal("value", macro.Steps[16].Property);
        Assert.Equal("Done", macro.Steps[16].Expected);
        Assert.Equal("Must be Done to continue", macro.Steps[16].Message);
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

    // --- Launch & WaitForWindow step parsing ---

    [Fact]
    public void Load_LaunchStep_ParsesAllFields()
    {
        WriteMacro("launch-test.yaml", @"
name: Launch Test
description: Tests launch step parsing
steps:
  - action: launch
    exe_path: ""C:\\Program Files\\MyApp\\App.exe""
    arguments: --config default
    working_directory: ""C:\\Program Files\\MyApp""
    if_not_running: true
    timeout: 45
");

        using var engine = new MacroEngine(_tempDir, enableWatcher: false);
        var macro = engine.Get("launch-test");
        Assert.NotNull(macro);
        var step = macro!.Steps[0];
        Assert.Equal("launch", step.Action);
        Assert.Equal("C:\\Program Files\\MyApp\\App.exe", step.ExePath);
        Assert.Equal("--config default", step.Arguments);
        Assert.Equal("C:\\Program Files\\MyApp", step.WorkingDirectory);
        Assert.True(step.IfNotRunning);
        Assert.Equal(45, step.StepTimeout);
    }

    [Fact]
    public void Load_LaunchStep_IfNotRunningDefaultsNull()
    {
        WriteMacro("launch-default.yaml", @"
name: Launch Default
description: Tests if_not_running default
steps:
  - action: launch
    exe_path: app.exe
");

        using var engine = new MacroEngine(_tempDir, enableWatcher: false);
        var macro = engine.Get("launch-default");
        Assert.NotNull(macro);
        var step = macro!.Steps[0];
        Assert.Equal("launch", step.Action);
        Assert.Null(step.IfNotRunning); // null → engine defaults to true
    }

    [Fact]
    public void Load_WaitForWindowStep_ParsesAllFields()
    {
        WriteMacro("wait-win.yaml", @"
name: Wait Window
description: Tests wait_for_window parsing
steps:
  - action: wait_for_window
    title_contains: My Application
    automation_id: mainPanel
    name: Main Panel
    control_type: Custom
    timeout: 30
    retry_interval: 2
");

        using var engine = new MacroEngine(_tempDir, enableWatcher: false);
        var macro = engine.Get("wait-win");
        Assert.NotNull(macro);
        var step = macro!.Steps[0];
        Assert.Equal("wait_for_window", step.Action);
        Assert.Equal("My Application", step.TitleContains);
        Assert.Equal("mainPanel", step.AutomationId);
        Assert.Equal("Main Panel", step.Name);
        Assert.Equal("Custom", step.ControlType);
        Assert.Equal(30, step.StepTimeout);
        Assert.Equal(2, step.RetryInterval);
    }

    [Fact]
    public void Load_LaunchMacroWithParameters_SubstitutesExePath()
    {
        WriteMacro("launch-param.yaml", @"
name: Launch Param
description: Tests parameter substitution in exe_path
parameters:
  - name: exePath
    required: true
steps:
  - action: launch
    exe_path: ""{{exePath}}""
    if_not_running: true
");

        using var engine = new MacroEngine(_tempDir, enableWatcher: false);
        var macro = engine.Get("launch-param");
        Assert.NotNull(macro);
        Assert.Equal("{{exePath}}", macro!.Steps[0].ExePath);
        // Verify SubstituteParams works with exe_path value
        var substituted = MacroEngine.SubstituteParams(
            macro.Steps[0].ExePath,
            new Dictionary<string, string> { ["exePath"] = "C:\\MyApp.exe" });
        Assert.Equal("C:\\MyApp.exe", substituted);
    }

    [Fact]
    public async Task Execute_LaunchStepWithoutExePath_ReturnsError()
    {
        WriteMacro("launch-nope.yaml", @"
name: Launch No Path
description: Missing exe_path
steps:
  - action: launch
");

        using var engine = new MacroEngine(_tempDir, enableWatcher: false);
        var result = await engine.ExecuteAsync("launch-nope");
        Assert.False(result.Success);
        Assert.Contains("exe_path", result.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Load_FullLaunchMacro_ParsesLaunchAndWaitSteps()
    {
        WriteMacro(Path.Combine("product", "launch.yaml"), @"
name: Launch Product
description: Full launch macro with both steps
timeout: 60
parameters:
  - name: exePath
    description: Path to exe
    required: false
    default: ""C:\\App\\App.exe""
steps:
  - action: launch
    exe_path: ""{{exePath}}""
    if_not_running: true
    timeout: 45
  - action: wait_for_window
    title_contains: My Product
    timeout: 45
    retry_interval: 3
");

        using var engine = new MacroEngine(_tempDir, enableWatcher: false);
        var macros = engine.List();
        Assert.Single(macros);
        Assert.Equal("product/launch", macros[0].Name);
        var macro = engine.Get("product/launch");
        Assert.NotNull(macro);
        Assert.Equal(2, macro!.Steps.Count);
        Assert.Equal("launch", macro.Steps[0].Action);
        Assert.Equal("wait_for_window", macro.Steps[1].Action);
        Assert.Equal("My Product", macro.Steps[1].TitleContains);
    }

    [Fact]
    public void Load_WorkflowMacroCallingLaunch_ParsesNestedMacroStep()
    {
        WriteMacro(Path.Combine("product", "launch.yaml"), @"
name: Launch
description: Launch the product
steps:
  - action: launch
    exe_path: app.exe
    if_not_running: true
  - action: wait_for_window
    title_contains: Product
    timeout: 30
");

        WriteMacro(Path.Combine("product", "do-work.yaml"), @"
name: Do Work
description: A workflow that calls launch first
timeout: 90
parameters:
  - name: filePath
    required: true
steps:
  - action: macro
    macro_name: product/launch
  - action: focus
  - action: type
    text: ""{{filePath}}""
");

        using var engine = new MacroEngine(_tempDir, enableWatcher: false);
        var macros = engine.List();
        Assert.Equal(2, macros.Count);
        var workflow = engine.Get("product/do-work");
        Assert.NotNull(workflow);
        Assert.Equal(3, workflow!.Steps.Count);
        Assert.Equal("macro", workflow.Steps[0].Action);
        Assert.Equal("product/launch", workflow.Steps[0].MacroName);
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

    // --- ValidateSteps ---

    [Fact]
    public void ValidateSteps_EmptyList_ReturnsError()
    {
        var error = MacroEngine.ValidateSteps(new List<Dictionary<string, object>>());
        Assert.NotNull(error);
        Assert.Contains("at least one step", error);
    }

    [Fact]
    public void ValidateSteps_MissingAction_ReturnsError()
    {
        var steps = new List<Dictionary<string, object>>
        {
            new() { ["automation_id"] = "myBtn" }
        };
        var error = MacroEngine.ValidateSteps(steps);
        Assert.NotNull(error);
        Assert.Contains("missing 'action'", error);
    }

    [Fact]
    public void ValidateSteps_UnknownAction_ReturnsError()
    {
        var steps = new List<Dictionary<string, object>>
        {
            new() { ["action"] = "dance" }
        };
        var error = MacroEngine.ValidateSteps(steps);
        Assert.NotNull(error);
        Assert.Contains("unknown action 'dance'", error);
        Assert.Contains("Valid actions:", error);
    }

    [Fact]
    public void ValidateSteps_SendKeysWithoutKeys_ReturnsError()
    {
        var steps = new List<Dictionary<string, object>>
        {
            new() { ["action"] = "send_keys" }
        };
        var error = MacroEngine.ValidateSteps(steps);
        Assert.NotNull(error);
        Assert.Contains("requires 'keys'", error);
    }

    [Fact]
    public void ValidateSteps_FindWithoutSearchProps_ReturnsError()
    {
        var steps = new List<Dictionary<string, object>>
        {
            new() { ["action"] = "find" }
        };
        var error = MacroEngine.ValidateSteps(steps);
        Assert.NotNull(error);
        Assert.Contains("at least one of", error);
    }

    [Fact]
    public void ValidateSteps_TypeWithoutText_ReturnsError()
    {
        var steps = new List<Dictionary<string, object>>
        {
            new() { ["action"] = "type" }
        };
        var error = MacroEngine.ValidateSteps(steps);
        Assert.NotNull(error);
        Assert.Contains("requires 'text'", error);
    }

    [Fact]
    public void ValidateSteps_WaitWithoutSeconds_ReturnsError()
    {
        var steps = new List<Dictionary<string, object>>
        {
            new() { ["action"] = "wait" }
        };
        var error = MacroEngine.ValidateSteps(steps);
        Assert.NotNull(error);
        Assert.Contains("requires 'seconds'", error);
    }

    [Fact]
    public void ValidateSteps_MacroWithoutMacroName_ReturnsError()
    {
        var steps = new List<Dictionary<string, object>>
        {
            new() { ["action"] = "macro" }
        };
        var error = MacroEngine.ValidateSteps(steps);
        Assert.NotNull(error);
        Assert.Contains("requires 'macro_name'", error);
    }

    [Fact]
    public void ValidateSteps_WaitForWindowWithoutTitle_ReturnsError()
    {
        var steps = new List<Dictionary<string, object>>
        {
            new() { ["action"] = "wait_for_window" }
        };
        var error = MacroEngine.ValidateSteps(steps);
        Assert.NotNull(error);
        Assert.Contains("requires 'title_contains'", error);
    }

    [Fact]
    public void ValidateSteps_SetValueWithoutRef_ReturnsError()
    {
        var steps = new List<Dictionary<string, object>>
        {
            new() { ["action"] = "set_value", ["value"] = "hello" }
        };
        var error = MacroEngine.ValidateSteps(steps);
        Assert.NotNull(error);
        Assert.Contains("requires 'ref'", error);
    }

    [Fact]
    public void ValidateSteps_SetValueWithoutValue_ReturnsError()
    {
        var steps = new List<Dictionary<string, object>>
        {
            new() { ["action"] = "set_value", ["ref"] = "e1" }
        };
        var error = MacroEngine.ValidateSteps(steps);
        Assert.NotNull(error);
        Assert.Contains("requires 'value'", error);
    }

    [Fact]
    public void ValidateSteps_FindByPathWithoutPath_ReturnsError()
    {
        var steps = new List<Dictionary<string, object>>
        {
            new() { ["action"] = "find_by_path" }
        };
        var error = MacroEngine.ValidateSteps(steps);
        Assert.NotNull(error);
        Assert.Contains("requires 'path'", error);
    }

    [Fact]
    public void ValidateSteps_AttachWithoutProcessNameOrPid_ReturnsError()
    {
        var steps = new List<Dictionary<string, object>>
        {
            new() { ["action"] = "attach" }
        };
        var error = MacroEngine.ValidateSteps(steps);
        Assert.NotNull(error);
        Assert.Contains("at least one of", error);
    }

    [Fact]
    public void ValidateSteps_FileDialogWithoutText_ReturnsError()
    {
        var steps = new List<Dictionary<string, object>>
        {
            new() { ["action"] = "file_dialog" }
        };
        var error = MacroEngine.ValidateSteps(steps);
        Assert.NotNull(error);
        Assert.Contains("requires 'text'", error);
    }

    [Fact]
    public void ValidateSteps_ValidMultiStep_ReturnsNull()
    {
        var steps = new List<Dictionary<string, object>>
        {
            new() { ["action"] = "find", ["automation_id"] = "myBtn", ["save_as"] = "btn" },
            new() { ["action"] = "click", ["ref"] = "btn" },
            new() { ["action"] = "send_keys", ["keys"] = "Ctrl+S" },
            new() { ["action"] = "wait", ["seconds"] = 2 },
            new() { ["action"] = "focus" },
            new() { ["action"] = "snapshot" },
            new() { ["action"] = "screenshot" },
            new() { ["action"] = "verify", ["ref"] = "btn", ["property"] = "value", ["expected"] = "Done" },
        };
        var error = MacroEngine.ValidateSteps(steps);
        Assert.Null(error);
    }

    [Fact]
    public void ValidateSteps_NoRequiredFieldActions_ReturnsNull()
    {
        // Actions with no strictly required fields
        var actions = new[] { "click", "right_click", "focus", "snapshot", "screenshot", "properties", "children", "launch" };
        foreach (var action in actions)
        {
            var steps = new List<Dictionary<string, object>>
            {
                new() { ["action"] = action }
            };
            var error = MacroEngine.ValidateSteps(steps);
            Assert.Null(error);
        }
    }

    [Fact]
    public void ValidateSteps_KeysAlias_WorksLikeSendKeys()
    {
        // "keys" is an alias for "send_keys"
        var steps = new List<Dictionary<string, object>>
        {
            new() { ["action"] = "keys", ["keys"] = "Ctrl+S" }
        };
        var error = MacroEngine.ValidateSteps(steps);
        Assert.Null(error);

        // But missing the keys field should fail
        var steps2 = new List<Dictionary<string, object>>
        {
            new() { ["action"] = "keys" }
        };
        var error2 = MacroEngine.ValidateSteps(steps2);
        Assert.NotNull(error2);
        Assert.Contains("requires 'keys'", error2);
    }

    [Fact]
    public void ValidateSteps_ErrorReportsCorrectStepNumber()
    {
        var steps = new List<Dictionary<string, object>>
        {
            new() { ["action"] = "focus" },
            new() { ["action"] = "find", ["automation_id"] = "ok" },
            new() { ["action"] = "type" }, // step 3: missing text
        };
        var error = MacroEngine.ValidateSteps(steps);
        Assert.NotNull(error);
        Assert.StartsWith("Step 3", error);
    }

    // --- wait_for_enabled validation ---

    [Fact]
    public void ValidateSteps_WaitForEnabledWithAutomationId_ReturnsNull()
    {
        var steps = new List<Dictionary<string, object>>
        {
            new() { ["action"] = "wait_for_enabled", ["automation_id"] = "uxRibbons" }
        };
        var error = MacroEngine.ValidateSteps(steps);
        Assert.Null(error);
    }

    [Fact]
    public void ValidateSteps_WaitForEnabledWithRef_ReturnsNull()
    {
        var steps = new List<Dictionary<string, object>>
        {
            new() { ["action"] = "wait_for_enabled", ["ref"] = "e1" }
        };
        var error = MacroEngine.ValidateSteps(steps);
        Assert.Null(error);
    }

    [Fact]
    public void ValidateSteps_WaitForEnabledWithName_ReturnsNull()
    {
        var steps = new List<Dictionary<string, object>>
        {
            new() { ["action"] = "wait_for_enabled", ["name"] = "S2 // Diagnostics" }
        };
        var error = MacroEngine.ValidateSteps(steps);
        Assert.Null(error);
    }

    [Fact]
    public void ValidateSteps_WaitForEnabledWithoutAnyCriteria_ReturnsError()
    {
        var steps = new List<Dictionary<string, object>>
        {
            new() { ["action"] = "wait_for_enabled" }
        };
        var error = MacroEngine.ValidateSteps(steps);
        Assert.NotNull(error);
        Assert.Contains("at least one of", error);
    }

    [Fact]
    public void Load_WaitForEnabledStep_ParsesAllFields()
    {
        WriteMacro("wait-enabled.yaml", @"
name: Wait Enabled
description: Tests wait_for_enabled parsing
steps:
  - action: wait_for_enabled
    automation_id: uxRibbons
    enabled: true
    timeout: 30
    retry_interval: 0.5
    save_as: tab
");

        using var engine = new MacroEngine(_tempDir, enableWatcher: false);
        var macro = engine.Get("wait-enabled");
        Assert.NotNull(macro);
        var step = macro!.Steps[0];
        Assert.Equal("wait_for_enabled", step.Action);
        Assert.Equal("uxRibbons", step.AutomationId);
        Assert.True(step.Enabled);
        Assert.Equal(30, step.StepTimeout);
        Assert.Equal(0.5, step.RetryInterval);
        Assert.Equal("tab", step.SaveAs);
    }

    [Fact]
    public void Load_WaitForEnabledStep_EnabledFalse_ParsesCorrectly()
    {
        WriteMacro("wait-disabled.yaml", @"
name: Wait Disabled
description: Tests wait_for_enabled with enabled=false
steps:
  - action: wait_for_enabled
    automation_id: uxRibbons
    enabled: false
    timeout: 10
");

        using var engine = new MacroEngine(_tempDir, enableWatcher: false);
        var macro = engine.Get("wait-disabled");
        Assert.NotNull(macro);
        var step = macro!.Steps[0];
        Assert.Equal("wait_for_enabled", step.Action);
        Assert.False(step.Enabled);
        Assert.Equal(10, step.StepTimeout);
    }

    [Fact]
    public void Load_WaitForEnabledStep_EnabledOmitted_DefaultsToNull()
    {
        WriteMacro("wait-default.yaml", @"
name: Wait Default
description: Tests wait_for_enabled with enabled omitted
steps:
  - action: wait_for_enabled
    automation_id: uxRibbons
    timeout: 30
");

        using var engine = new MacroEngine(_tempDir, enableWatcher: false);
        var macro = engine.Get("wait-default");
        Assert.NotNull(macro);
        var step = macro!.Steps[0];
        Assert.Null(step.Enabled); // defaults to null, runtime treats as true
    }

    // --- verify step parsing ---

    [Fact]
    public void Load_VerifyStep_ParsesAllFields()
    {
        WriteMacro("verify-test.yaml", @"
name: Verify Test
description: Tests verify step parsing
steps:
  - action: find
    automation_id: uxStatus
    save_as: status
  - action: verify
    ref: status
    property: value
    expected: Complete
    message: Status must be Complete before proceeding
");

        using var engine = new MacroEngine(_tempDir, enableWatcher: false);
        var macro = engine.Get("verify-test");
        Assert.NotNull(macro);
        Assert.Equal(2, macro!.Steps.Count);
        var step = macro.Steps[1];
        Assert.Equal("verify", step.Action);
        Assert.Equal("status", step.Ref);
        Assert.Equal("value", step.Property);
        Assert.Equal("Complete", step.Expected);
        Assert.Equal("Status must be Complete before proceeding", step.Message);
    }

    [Fact]
    public void Load_VerifyStep_MinimalFields_ParsesCorrectly()
    {
        WriteMacro("verify-minimal.yaml", @"
name: Verify Minimal
description: Tests verify with minimal fields
steps:
  - action: verify
    ref: el
    property: is_enabled
    expected: 'True'
");

        using var engine = new MacroEngine(_tempDir, enableWatcher: false);
        var macro = engine.Get("verify-minimal");
        Assert.NotNull(macro);
        var step = macro!.Steps[0];
        Assert.Equal("verify", step.Action);
        Assert.Equal("el", step.Ref);
        Assert.Equal("is_enabled", step.Property);
        Assert.Equal("True", step.Expected);
        Assert.Null(step.Message); // message is optional
    }

    [Fact]
    public void ValidateSteps_VerifyWithoutRef_ReturnsError()
    {
        var steps = new List<Dictionary<string, object>>
        {
            new() { ["action"] = "verify", ["property"] = "value", ["expected"] = "Done" }
        };
        var error = MacroEngine.ValidateSteps(steps);
        Assert.NotNull(error);
        Assert.Contains("requires 'ref'", error);
    }

    [Fact]
    public void ValidateSteps_VerifyWithoutProperty_ReturnsError()
    {
        var steps = new List<Dictionary<string, object>>
        {
            new() { ["action"] = "verify", ["ref"] = "el", ["expected"] = "Done" }
        };
        var error = MacroEngine.ValidateSteps(steps);
        Assert.NotNull(error);
        Assert.Contains("requires 'property'", error);
    }

    [Fact]
    public void ValidateSteps_VerifyWithoutExpected_ReturnsError()
    {
        var steps = new List<Dictionary<string, object>>
        {
            new() { ["action"] = "verify", ["ref"] = "el", ["property"] = "value" }
        };
        var error = MacroEngine.ValidateSteps(steps);
        Assert.NotNull(error);
        Assert.Contains("requires 'expected'", error);
    }

    [Fact]
    public void ValidateSteps_VerifyWithAllFields_ReturnsNull()
    {
        var steps = new List<Dictionary<string, object>>
        {
            new() { ["action"] = "verify", ["ref"] = "el", ["property"] = "value", ["expected"] = "Done" }
        };
        var error = MacroEngine.ValidateSteps(steps);
        Assert.Null(error);
    }

    [Fact]
    public void ValidateSteps_VerifyWithMessage_ReturnsNull()
    {
        var steps = new List<Dictionary<string, object>>
        {
            new() { ["action"] = "verify", ["ref"] = "el", ["property"] = "toggle_state", ["expected"] = "On", ["message"] = "Checkbox must be checked" }
        };
        var error = MacroEngine.ValidateSteps(steps);
        Assert.Null(error);
    }

    // --- GetProductFolder ---

    [Fact]
    public void GetProductFolder_MatchingProcessName_ReturnsFolder()
    {
        WriteKnowledgeBase("my-app", "MyApp");

        using var engine = new MacroEngine(_tempDir, enableWatcher: false);
        var folder = engine.GetProductFolder("MyApp");
        Assert.Equal("my-app", folder);
    }

    [Fact]
    public void GetProductFolder_CaseInsensitive_ReturnsFolder()
    {
        WriteKnowledgeBase("my-app", "MyApp");

        using var engine = new MacroEngine(_tempDir, enableWatcher: false);
        var folder = engine.GetProductFolder("myapp");
        Assert.Equal("my-app", folder);
    }

    [Fact]
    public void GetProductFolder_NoMatch_ReturnsNull()
    {
        WriteKnowledgeBase("my-app", "MyApp");

        using var engine = new MacroEngine(_tempDir, enableWatcher: false);
        var folder = engine.GetProductFolder("SomeOtherApp");
        Assert.Null(folder);
    }

    [Fact]
    public void GetProductFolder_NoKnowledgeBases_ReturnsNull()
    {
        using var engine = new MacroEngine(_tempDir, enableWatcher: false);
        var folder = engine.GetProductFolder("Anything");
        Assert.Null(folder);
    }

    // --- SaveMacro ---

    [Fact]
    public void SaveMacro_ValidSteps_CreatesFile()
    {
        WriteKnowledgeBase("test-app", "TestApp");

        using var engine = new MacroEngine(_tempDir, enableWatcher: false);

        var steps = new List<Dictionary<string, object>>
        {
            new() { ["action"] = "find", ["automation_id"] = "myBtn", ["save_as"] = "btn" },
            new() { ["action"] = "click", ["ref"] = "btn" },
        };

        var result = engine.SaveMacro("my-workflow", "A test workflow", steps, null, 30, false, "TestApp");

        Assert.True(result.Ok);
        Assert.Equal("test-app/my-workflow", result.MacroName);
        Assert.True(File.Exists(result.FilePath));

        // Verify the YAML content
        var yaml = File.ReadAllText(result.FilePath);
        Assert.Contains("my-workflow", yaml);
        Assert.Contains("A test workflow", yaml);
        Assert.Contains("find", yaml);
        Assert.Contains("myBtn", yaml);
        Assert.Contains("click", yaml);
    }

    [Fact]
    public void SaveMacro_ExistingMacro_WithoutForce_ReturnsError()
    {
        WriteKnowledgeBase("test-app", "TestApp");

        using var engine = new MacroEngine(_tempDir, enableWatcher: false);

        var steps = new List<Dictionary<string, object>>
        {
            new() { ["action"] = "focus" },
        };

        // Save the first time
        var result1 = engine.SaveMacro("my-workflow", "First", steps, null, 30, false, "TestApp");
        Assert.True(result1.Ok);

        // Try to save again without force
        var result2 = engine.SaveMacro("my-workflow", "Second", steps, null, 30, false, "TestApp");
        Assert.False(result2.Ok);
        Assert.Contains("already exists", result2.Message);
    }

    [Fact]
    public void SaveMacro_ExistingMacro_WithForce_Overwrites()
    {
        WriteKnowledgeBase("test-app", "TestApp");

        using var engine = new MacroEngine(_tempDir, enableWatcher: false);

        var steps = new List<Dictionary<string, object>>
        {
            new() { ["action"] = "focus" },
        };

        engine.SaveMacro("my-workflow", "First", steps, null, 30, false, "TestApp");
        var result = engine.SaveMacro("my-workflow", "Updated", steps, null, 30, true, "TestApp");

        Assert.True(result.Ok);
        var yaml = File.ReadAllText(result.FilePath);
        Assert.Contains("Updated", yaml);
    }

    [Fact]
    public void SaveMacro_InvalidSteps_ReturnsError()
    {
        WriteKnowledgeBase("test-app", "TestApp");

        using var engine = new MacroEngine(_tempDir, enableWatcher: false);

        var steps = new List<Dictionary<string, object>>
        {
            new() { ["action"] = "type" }, // missing text
        };

        var result = engine.SaveMacro("bad-macro", "Bad", steps, null, 30, false, "TestApp");
        Assert.False(result.Ok);
        Assert.Contains("Validation error", result.Message);
    }

    [Fact]
    public void SaveMacro_NoMatchingKnowledgeBase_ReturnsError()
    {
        using var engine = new MacroEngine(_tempDir, enableWatcher: false);

        var steps = new List<Dictionary<string, object>>
        {
            new() { ["action"] = "focus" },
        };

        var result = engine.SaveMacro("my-macro", "Test", steps, null, 30, false, "UnknownApp");
        Assert.False(result.Ok);
        Assert.Contains("Cannot determine product folder", result.Message);
    }

    [Fact]
    public void SaveMacro_NameWithProductPrefix_UsesAsIs()
    {
        WriteKnowledgeBase("test-app", "TestApp");

        using var engine = new MacroEngine(_tempDir, enableWatcher: false);

        var steps = new List<Dictionary<string, object>>
        {
            new() { ["action"] = "focus" },
        };

        // Name already includes product folder
        var result = engine.SaveMacro("test-app/custom-workflow", "Custom", steps, null, 30, false, "TestApp");
        Assert.True(result.Ok);
        Assert.Equal("test-app/custom-workflow", result.MacroName);
        Assert.True(File.Exists(result.FilePath));
    }

    [Fact]
    public void SaveMacro_WithParameters_IncludesInYaml()
    {
        WriteKnowledgeBase("test-app", "TestApp");

        using var engine = new MacroEngine(_tempDir, enableWatcher: false);

        var steps = new List<Dictionary<string, object>>
        {
            new() { ["action"] = "type", ["text"] = "{{filePath}}" },
        };

        var parameters = new List<Dictionary<string, object>>
        {
            new()
            {
                ["name"] = "filePath",
                ["description"] = "Path to the file",
                ["required"] = true,
                ["default"] = "C:\\test.xer"
            }
        };

        var result = engine.SaveMacro("param-macro", "With params", steps, parameters, 30, false, "TestApp");
        Assert.True(result.Ok);

        var yaml = File.ReadAllText(result.FilePath);
        Assert.Contains("filePath", yaml);
        Assert.Contains("Path to the file", yaml);
    }

    [Fact]
    public void SaveMacro_WithTimeout_IncludesInYaml()
    {
        WriteKnowledgeBase("test-app", "TestApp");

        using var engine = new MacroEngine(_tempDir, enableWatcher: false);

        var steps = new List<Dictionary<string, object>>
        {
            new() { ["action"] = "focus" },
        };

        var result = engine.SaveMacro("timeout-macro", "With timeout", steps, null, 60, false, "TestApp");
        Assert.True(result.Ok);

        var yaml = File.ReadAllText(result.FilePath);
        Assert.Contains("timeout: 60", yaml);
    }

    [Fact]
    public void SaveMacro_ZeroTimeout_OmitsFromYaml()
    {
        WriteKnowledgeBase("test-app", "TestApp");

        using var engine = new MacroEngine(_tempDir, enableWatcher: false);

        var steps = new List<Dictionary<string, object>>
        {
            new() { ["action"] = "focus" },
        };

        var result = engine.SaveMacro("no-timeout", "No timeout", steps, null, 0, false, "TestApp");
        Assert.True(result.Ok);

        var yaml = File.ReadAllText(result.FilePath);
        Assert.DoesNotContain("timeout:", yaml);
    }

    [Fact]
    public void MacrosPath_ReturnsConfiguredPath()
    {
        using var engine = new MacroEngine(_tempDir, enableWatcher: false);
        Assert.Equal(_tempDir, engine.MacrosPath);
    }

    // --- Helper to write a knowledge base ---

    private void WriteKnowledgeBase(string productFolder, string processName)
    {
        WriteMacro(Path.Combine(productFolder, "_knowledge.yaml"), $@"
kind: knowledge-base

application:
  name: {productFolder}
  process_name: {processName}

workflows: {{}}
");
    }
}
