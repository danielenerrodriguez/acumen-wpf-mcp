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
    match_mode: contains
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
        Assert.Equal("contains", macro.Steps[16].MatchMode);
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

    // --- verify match_mode ---

    [Fact]
    public void Load_VerifyStep_MatchMode_ParsesCorrectly()
    {
        WriteMacro("verify-matchmode.yaml", @"
name: Verify MatchMode
description: Tests verify match_mode parsing
steps:
  - action: verify
    ref: el
    property: name
    expected: 'Workbook1 (0)'
    match_mode: not_equals
    message: Import failed
");

        using var engine = new MacroEngine(_tempDir, enableWatcher: false);
        var macro = engine.Get("verify-matchmode");
        Assert.NotNull(macro);
        var step = macro!.Steps[0];
        Assert.Equal("verify", step.Action);
        Assert.Equal("el", step.Ref);
        Assert.Equal("name", step.Property);
        Assert.Equal("Workbook1 (0)", step.Expected);
        Assert.Equal("not_equals", step.MatchMode);
        Assert.Equal("Import failed", step.Message);
    }

    [Fact]
    public void Load_VerifyStep_NoMatchMode_DefaultsToNull()
    {
        WriteMacro("verify-no-matchmode.yaml", @"
name: Verify No MatchMode
description: Tests verify without match_mode
steps:
  - action: verify
    ref: el
    property: value
    expected: Done
");

        using var engine = new MacroEngine(_tempDir, enableWatcher: false);
        var macro = engine.Get("verify-no-matchmode");
        Assert.NotNull(macro);
        Assert.Null(macro!.Steps[0].MatchMode);
    }

    [Fact]
    public void ValidateSteps_VerifyWithMatchMode_ReturnsNull()
    {
        var steps = new List<Dictionary<string, object>>
        {
            new() { ["action"] = "verify", ["ref"] = "el", ["property"] = "name", ["expected"] = "test", ["match_mode"] = "contains" }
        };
        var error = MacroEngine.ValidateSteps(steps);
        Assert.Null(error);
    }

    [Theory]
    [InlineData("hello world", "hello world", "equals", true)]
    [InlineData("Hello World", "hello world", "equals", true)]
    [InlineData("hello", "world", "equals", false)]
    [InlineData("hello world", "world", "contains", true)]
    [InlineData("Hello World", "WORLD", "contains", true)]
    [InlineData("hello", "xyz", "contains", false)]
    [InlineData("hello", "world", "not_equals", true)]
    [InlineData("hello", "hello", "not_equals", false)]
    [InlineData("Hello", "HELLO", "not_equals", false)]
    [InlineData("Workbook1 (56)", @"\(\d+\)", "regex", true)]
    [InlineData("Workbook1 (0)", @"\(\d*[1-9]\d*\)", "regex", false)]
    [InlineData("hello world", "hello", "starts_with", true)]
    [InlineData("Hello World", "HELLO", "starts_with", true)]
    [InlineData("hello world", "world", "starts_with", false)]
    public void VerifyMatch_AllModes_ReturnsExpected(string actual, string expected, string matchMode, bool expectedResult)
    {
        var result = MacroEngine.VerifyMatch(actual, expected, matchMode);
        Assert.NotNull(result);
        Assert.Equal(expectedResult, result!.Value);
    }

    [Fact]
    public void VerifyMatch_UnknownMode_ReturnsNull()
    {
        var result = MacroEngine.VerifyMatch("hello", "hello", "invalid_mode");
        Assert.Null(result);
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

    // --- Include step expansion tests ---

    [Fact]
    public void Include_BasicExpansion_FlattensSteps()
    {
        // Child macro with 2 steps
        WriteMacro("child.yaml", @"
name: Child
description: A reusable child macro
steps:
  - action: focus
  - action: send_keys
    keys: Alt,F
");
        // Parent macro includes child
        WriteMacro("parent.yaml", @"
name: Parent
description: Uses include
steps:
  - action: wait
    seconds: 1
  - action: include
    macro_name: child
  - action: send_keys
    keys: Enter
");
        using var engine = new MacroEngine(_tempDir, enableWatcher: false);
        var parent = engine.Get("parent");
        Assert.NotNull(parent);
        // Include step should be replaced with child's 2 steps: wait + focus + send_keys(Alt,F) + send_keys(Enter)
        Assert.Equal(4, parent!.Steps.Count);
        Assert.Equal("wait", parent.Steps[0].Action);
        Assert.Equal("focus", parent.Steps[1].Action);
        Assert.Equal("send_keys", parent.Steps[2].Action);
        Assert.Equal("Alt,F", parent.Steps[2].Keys);
        Assert.Equal("send_keys", parent.Steps[3].Action);
        Assert.Equal("Enter", parent.Steps[3].Keys);
    }

    [Fact]
    public void Include_ParameterMapping_RemapsChildParams()
    {
        WriteMacro("child.yaml", @"
name: Child
parameters:
  - name: filePath
    required: true
steps:
  - action: type
    text: '{{filePath}}'
");
        // Parent uses a different param name and maps it
        WriteMacro("parent.yaml", @"
name: Parent
parameters:
  - name: inputFile
    required: true
steps:
  - action: include
    macro_name: child
    params:
      filePath: '{{inputFile}}'
");
        using var engine = new MacroEngine(_tempDir, enableWatcher: false);
        var parent = engine.Get("parent");
        Assert.NotNull(parent);
        Assert.Single(parent!.Steps);
        Assert.Equal("type", parent.Steps[0].Action);
        // {{filePath}} in child should now be {{inputFile}} in parent
        Assert.Equal("{{inputFile}}", parent.Steps[0].Text);
    }

    [Fact]
    public void Include_LiteralParamValue_SubstitutedDirectly()
    {
        WriteMacro("child.yaml", @"
name: Child
parameters:
  - name: filePath
    required: true
steps:
  - action: type
    text: '{{filePath}}'
");
        WriteMacro("parent.yaml", @"
name: Parent
steps:
  - action: include
    macro_name: child
    params:
      filePath: 'C:\data\test.xer'
");
        using var engine = new MacroEngine(_tempDir, enableWatcher: false);
        var parent = engine.Get("parent");
        Assert.NotNull(parent);
        Assert.Single(parent!.Steps);
        // Literal value should replace {{filePath}} directly
        Assert.Equal(@"C:\data\test.xer", parent.Steps[0].Text);
    }

    [Fact]
    public void Include_CircularDetection_ReportsLoadError()
    {
        WriteMacro("a.yaml", @"
name: A
steps:
  - action: include
    macro_name: b
");
        WriteMacro("b.yaml", @"
name: B
steps:
  - action: include
    macro_name: a
");
        using var engine = new MacroEngine(_tempDir, enableWatcher: false);
        Assert.NotEmpty(engine.LoadErrors);
        Assert.Contains(engine.LoadErrors, e => e.Error.Contains("Circular include"));
    }

    [Fact]
    public void Include_NonexistentMacro_ReportsLoadError()
    {
        WriteMacro("parent.yaml", @"
name: Parent
steps:
  - action: include
    macro_name: does-not-exist
");
        using var engine = new MacroEngine(_tempDir, enableWatcher: false);
        Assert.NotEmpty(engine.LoadErrors);
        Assert.Contains(engine.LoadErrors, e => e.Error.Contains("unknown macro 'does-not-exist'"));
    }

    [Fact]
    public void Include_NestedIncludes_FullyFlattened()
    {
        WriteMacro("c.yaml", @"
name: C
steps:
  - action: focus
");
        WriteMacro("b.yaml", @"
name: B
steps:
  - action: include
    macro_name: c
  - action: send_keys
    keys: Alt,F
");
        WriteMacro("a.yaml", @"
name: A
steps:
  - action: wait
    seconds: 1
  - action: include
    macro_name: b
  - action: send_keys
    keys: Enter
");
        using var engine = new MacroEngine(_tempDir, enableWatcher: false);
        var a = engine.Get("a");
        Assert.NotNull(a);
        // A: wait + (B: (C: focus) + send_keys(Alt,F)) + send_keys(Enter) = 4 steps
        Assert.Equal(4, a!.Steps.Count);
        Assert.Equal("wait", a.Steps[0].Action);
        Assert.Equal("focus", a.Steps[1].Action);
        Assert.Equal("send_keys", a.Steps[2].Action);
        Assert.Equal("Alt,F", a.Steps[2].Keys);
        Assert.Equal("send_keys", a.Steps[3].Action);
        Assert.Equal("Enter", a.Steps[3].Keys);
    }

    [Fact]
    public void Include_MultipleIncludes_BothExpanded()
    {
        WriteMacro("launch-steps.yaml", @"
name: Launch Steps
steps:
  - action: launch
    exe_path: test.exe
    if_not_running: true
  - action: wait_for_window
    title_contains: Test
    timeout: 30
");
        WriteMacro("verify-steps.yaml", @"
name: Verify Steps
steps:
  - action: find
    automation_id: myElement
    save_as: el
  - action: verify
    ref: el
    property: name
    expected: Done
");
        WriteMacro("parent.yaml", @"
name: Parent
steps:
  - action: include
    macro_name: launch-steps
  - action: focus
  - action: include
    macro_name: verify-steps
");
        using var engine = new MacroEngine(_tempDir, enableWatcher: false);
        var parent = engine.Get("parent");
        Assert.NotNull(parent);
        // launch(2) + focus(1) + verify-steps(2) = 5
        Assert.Equal(5, parent!.Steps.Count);
        Assert.Equal("launch", parent.Steps[0].Action);
        Assert.Equal("wait_for_window", parent.Steps[1].Action);
        Assert.Equal("focus", parent.Steps[2].Action);
        Assert.Equal("find", parent.Steps[3].Action);
        Assert.Equal("verify", parent.Steps[4].Action);
    }

    [Fact]
    public void Include_PreservesOriginalMacroSteps()
    {
        // Verify that the included macro's own step list isn't mutated
        WriteMacro("child.yaml", @"
name: Child
parameters:
  - name: val
steps:
  - action: type
    text: '{{val}}'
");
        WriteMacro("parent.yaml", @"
name: Parent
steps:
  - action: include
    macro_name: child
    params:
      val: hello
");
        using var engine = new MacroEngine(_tempDir, enableWatcher: false);
        var child = engine.Get("child");
        Assert.NotNull(child);
        // Child's original step should still have {{val}}, not 'hello'
        Assert.Equal("{{val}}", child!.Steps[0].Text);
    }

    [Fact]
    public void Include_InSubfolder_ResolvesCorrectly()
    {
        WriteMacro("product/child.yaml", @"
name: Child
steps:
  - action: focus
  - action: send_keys
    keys: F9
");
        WriteMacro("product/parent.yaml", @"
name: Parent
steps:
  - action: include
    macro_name: product/child
  - action: wait
    seconds: 1
");
        using var engine = new MacroEngine(_tempDir, enableWatcher: false);
        var parent = engine.Get("product/parent");
        Assert.NotNull(parent);
        Assert.Equal(3, parent!.Steps.Count);
        Assert.Equal("focus", parent.Steps[0].Action);
        Assert.Equal("send_keys", parent.Steps[1].Action);
        Assert.Equal("wait", parent.Steps[2].Action);
    }

    [Fact]
    public void Include_MissingMacroName_ReportsLoadError()
    {
        WriteMacro("parent.yaml", @"
name: Parent
steps:
  - action: include
");
        using var engine = new MacroEngine(_tempDir, enableWatcher: false);
        Assert.NotEmpty(engine.LoadErrors);
        Assert.Contains(engine.LoadErrors, e => e.Error.Contains("missing 'macro_name'"));
    }

    [Fact]
    public void Include_ValidateSteps_RequiresMacroName()
    {
        var steps = new List<Dictionary<string, object>>
        {
            new() { ["action"] = "include" }
        };
        var error = MacroEngine.ValidateSteps(steps);
        Assert.NotNull(error);
        Assert.Contains("macro_name", error);
    }

    [Fact]
    public void Include_ValidateSteps_AcceptsValidInclude()
    {
        var steps = new List<Dictionary<string, object>>
        {
            new() { ["action"] = "include", ["macro_name"] = "some-macro" }
        };
        var error = MacroEngine.ValidateSteps(steps);
        Assert.Null(error);
    }

    [Fact]
    public void Include_FormatStepSummary_ShowsMacroName()
    {
        var step = new MacroStep { Action = "include", MacroName = "acumen-fuse/import-xer" };
        var summary = MacroEngine.FormatStepSummary(step, new Dictionary<string, string>());
        Assert.Contains("macro=acumen-fuse/import-xer", summary);
    }

    // --- Macro step fix tests ---

    [Fact]
    public async Task Macro_NestedStep_PassesOnLog()
    {
        // Nested macro: a simple focus step (will fail without engine but tests log flow)
        WriteMacro("child.yaml", @"
name: Child
timeout: 10
steps:
  - action: wait
    seconds: 0.1
");
        WriteMacro("parent.yaml", @"
name: Parent
timeout: 30
steps:
  - action: wait
    seconds: 0.1
  - action: macro
    macro_name: child
");
        using var engine = new MacroEngine(_tempDir, enableWatcher: false);
        var logMessages = new List<string>();
        var result = await engine.ExecuteAsync("parent",
            onLog: msg => logMessages.Add(msg));
        Assert.True(result.Success);
        // Should have parent step logs AND nested child step logs with prefix
        Assert.Contains(logMessages, m => m.Contains("Step 2 >") && m.Contains("Step 1/1"));
    }

    [Fact]
    public async Task Macro_NestedStep_UsesNestedMacroTimeout()
    {
        // Child macro has a 30s timeout — the macro step should NOT use the 5s default
        WriteMacro("child.yaml", @"
name: Child
timeout: 30
steps:
  - action: wait
    seconds: 0.1
");
        WriteMacro("parent.yaml", @"
name: Parent
timeout: 60
steps:
  - action: macro
    macro_name: child
");
        using var engine = new MacroEngine(_tempDir, enableWatcher: false);
        var result = await engine.ExecuteAsync("parent");
        // Should succeed — child's wait is only 0.1s, well within 30s timeout
        Assert.True(result.Success);
    }

    [Fact]
    public async Task Macro_NestedStep_NotFound_ReturnsError()
    {
        WriteMacro("parent.yaml", @"
name: Parent
steps:
  - action: macro
    macro_name: nonexistent
");
        using var engine = new MacroEngine(_tempDir, enableWatcher: false);
        var result = await engine.ExecuteAsync("parent");
        Assert.False(result.Success);
        Assert.Contains("not found", result.Message);
    }

    // --- run_script step type ---

    [Fact]
    public void ValidateSteps_RunScriptWithoutCommand_ReturnsError()
    {
        var steps = new List<Dictionary<string, object>>
        {
            new() { ["action"] = "run_script" }
        };
        var error = MacroEngine.ValidateSteps(steps);
        Assert.NotNull(error);
        Assert.Contains("requires 'command'", error);
    }

    [Fact]
    public void ValidateSteps_RunScriptWithCommand_ReturnsNull()
    {
        var steps = new List<Dictionary<string, object>>
        {
            new() { ["action"] = "run_script", ["command"] = "cmd.exe" }
        };
        var error = MacroEngine.ValidateSteps(steps);
        Assert.Null(error);
    }

    [Fact]
    public void Load_RunScriptStep_ParsesAllFields()
    {
        WriteMacro("run-script-test.yaml", @"
name: Run Script Test
description: Tests run_script step parsing
steps:
  - action: run_script
    command: powershell.exe
    arguments: -NoProfile -Command ""echo hi""
    working_directory: C:\temp
    save_output_as: scriptResult
    ignore_exit_code: true
    timeout: 30
");
        using var engine = new MacroEngine(_tempDir, enableWatcher: false);
        var macro = engine.Get("run-script-test");
        Assert.NotNull(macro);
        var step = macro!.Steps[0];
        Assert.Equal("run_script", step.Action);
        Assert.Equal("powershell.exe", step.Command);
        Assert.Contains("-NoProfile", step.Arguments);
        Assert.Equal(@"C:\temp", step.WorkingDirectory);
        Assert.Equal("scriptResult", step.SaveOutputAs);
        Assert.True(step.IgnoreExitCode);
        Assert.Equal(30, step.StepTimeout);
    }

    [Fact]
    public void Load_RunScriptStep_OptionalFieldsDefaultToNull()
    {
        WriteMacro("run-script-minimal.yaml", @"
name: Minimal
steps:
  - action: run_script
    command: cmd.exe
");
        using var engine = new MacroEngine(_tempDir, enableWatcher: false);
        var macro = engine.Get("run-script-minimal");
        Assert.NotNull(macro);
        var step = macro!.Steps[0];
        Assert.Equal("run_script", step.Action);
        Assert.Equal("cmd.exe", step.Command);
        Assert.Null(step.Arguments);
        Assert.Null(step.WorkingDirectory);
        Assert.Null(step.SaveOutputAs);
        Assert.Null(step.IgnoreExitCode);
    }

    [Fact]
    public void RunScript_FormatStepSummary_ShowsCommandOnly()
    {
        var step = new MacroStep
        {
            Action = "run_script",
            Command = "powershell.exe",
            Arguments = "-NoProfile -Command test"
        };
        var summary = MacroEngine.FormatStepSummary(step, new Dictionary<string, string>());
        Assert.Contains("command=powershell.exe", summary);
        // Arguments are intentionally excluded to keep logs readable
        Assert.DoesNotContain("args=", summary);
    }

    [Fact]
    public void RunScript_FormatStepSummary_ShowsSaveOutputAs()
    {
        var step = new MacroStep
        {
            Action = "run_script",
            Command = "cmd.exe",
            SaveOutputAs = "myOutput"
        };
        var summary = MacroEngine.FormatStepSummary(step, new Dictionary<string, string>());
        Assert.Contains("save_output_as=myOutput", summary);
    }

    [Fact]
    public void RunScript_FormatStepSummary_ShowsIgnoreExitCode()
    {
        var step = new MacroStep
        {
            Action = "run_script",
            Command = "cmd.exe",
            IgnoreExitCode = true
        };
        var summary = MacroEngine.FormatStepSummary(step, new Dictionary<string, string>());
        Assert.Contains("ignore_exit_code", summary);
    }

    [Fact]
    public void RunScript_FormatStepSummary_SubstitutesParams()
    {
        var step = new MacroStep
        {
            Action = "run_script",
            Command = "{{myCmd}}",
            Arguments = "{{myArgs}}"
        };
        var parameters = new Dictionary<string, string>
        {
            ["myCmd"] = "powershell.exe",
            ["myArgs"] = "-File test.ps1"
        };
        var summary = MacroEngine.FormatStepSummary(step, parameters);
        Assert.Contains("command=powershell.exe", summary);
        // Arguments are excluded from summary
        Assert.DoesNotContain("args=", summary);
    }

    [Fact]
    public void FormatStepSummary_UsesDescriptionWhenPresent()
    {
        var step = new MacroStep
        {
            Action = "run_script",
            Command = "powershell.exe",
            Arguments = "-NoProfile -Command \"some very long script...\"",
            Description = "Check if already installed"
        };
        var summary = MacroEngine.FormatStepSummary(step, new Dictionary<string, string>());
        Assert.Equal("run_script \u2014 Check if already installed", summary);
        Assert.DoesNotContain("powershell", summary);
    }

    [Fact]
    public void FormatStepSummary_DescriptionSubstitutesParams()
    {
        var step = new MacroStep
        {
            Action = "run_script",
            Command = "powershell.exe",
            Description = "Install version {{version}}"
        };
        var parameters = new Dictionary<string, string> { ["version"] = "8.12" };
        var summary = MacroEngine.FormatStepSummary(step, parameters);
        Assert.Equal("run_script \u2014 Install version 8.12", summary);
    }

    [Fact]
    public void FormatStepSummary_DescriptionWorksForAllActionTypes()
    {
        var step = new MacroStep
        {
            Action = "find",
            AutomationId = "someId",
            Description = "Find the main panel"
        };
        var summary = MacroEngine.FormatStepSummary(step, new Dictionary<string, string>());
        Assert.Equal("find \u2014 Find the main panel", summary);
        Assert.DoesNotContain("automation_id", summary);
    }

    [Fact]
    public void Include_RunScript_RemapsCommandParams()
    {
        WriteMacro("child.yaml", @"
name: Child
parameters:
  - name: scriptPath
steps:
  - action: run_script
    command: '{{scriptPath}}'
    arguments: --flag
");
        WriteMacro("parent.yaml", @"
name: Parent
parameters:
  - name: myScript
steps:
  - action: include
    macro_name: child
    params:
      scriptPath: '{{myScript}}'
");
        using var engine = new MacroEngine(_tempDir, enableWatcher: false);
        var parent = engine.Get("parent");
        Assert.NotNull(parent);
        Assert.Single(parent!.Steps);
        Assert.Equal("run_script", parent.Steps[0].Action);
        Assert.Equal("{{myScript}}", parent.Steps[0].Command);
    }

    [Fact]
    public async Task Execute_RunScript_SuccessfulCommand_ReturnsSuccess()
    {
        WriteMacro("echo-test.yaml", @"
name: Echo Test
steps:
  - action: run_script
    command: cmd.exe
    arguments: /c echo hello
");
        using var engine = new MacroEngine(_tempDir, enableWatcher: false);
        var logMessages = new List<string>();
        var result = await engine.ExecuteAsync("echo-test", onLog: msg => logMessages.Add(msg));
        Assert.True(result.Success, result.Message);
        Assert.Contains(logMessages, m => m.Contains("Exit code 0"));
    }

    [Fact]
    public async Task Execute_RunScript_MissingCommand_ReturnsError()
    {
        WriteMacro("no-cmd.yaml", @"
name: No Cmd
steps:
  - action: run_script
");
        using var engine = new MacroEngine(_tempDir, enableWatcher: false);
        var result = await engine.ExecuteAsync("no-cmd");
        Assert.False(result.Success);
        Assert.Contains("command", result.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public async Task Execute_RunScript_NonZeroExit_ReturnsError()
    {
        WriteMacro("fail-exit.yaml", @"
name: Fail Exit
steps:
  - action: run_script
    command: cmd.exe
    arguments: /c exit 1
");
        using var engine = new MacroEngine(_tempDir, enableWatcher: false);
        var result = await engine.ExecuteAsync("fail-exit");
        Assert.False(result.Success);
    }

    [Fact]
    public async Task Execute_RunScript_NonZeroExitWithIgnore_ReturnsSuccess()
    {
        WriteMacro("ignore-exit.yaml", @"
name: Ignore Exit
steps:
  - action: run_script
    command: cmd.exe
    arguments: /c exit 42
    ignore_exit_code: true
");
        using var engine = new MacroEngine(_tempDir, enableWatcher: false);
        var logMessages = new List<string>();
        var result = await engine.ExecuteAsync("ignore-exit", onLog: msg => logMessages.Add(msg));
        Assert.True(result.Success, result.Message);
        Assert.Contains(logMessages, m => m.Contains("Exit code 42"));
    }

    [Fact]
    public async Task Execute_RunScript_SaveOutputAs_CapturesStdout()
    {
        // Use two run_script steps: first captures output, second verifies
        // the parameter is available (by using it in a command that would fail if empty)
        WriteMacro("capture-test.yaml", @"
name: Capture Test
steps:
  - action: run_script
    command: cmd.exe
    arguments: /c echo test_value_123
    save_output_as: captured
  - action: run_script
    command: cmd.exe
    arguments: /c echo got={{captured}}
    save_output_as: verification
");
        using var engine = new MacroEngine(_tempDir, enableWatcher: false);
        var logMessages = new List<string>();
        var result = await engine.ExecuteAsync("capture-test", onLog: msg => logMessages.Add(msg));
        Assert.True(result.Success, result.Message);
        // Check that the captured value was logged
        Assert.Contains(logMessages, m => m.Contains("captured") && m.Contains("test_value_123"));
        // Check that the second step used the captured value (visible in the step summary log)
        Assert.Contains(logMessages, m => m.Contains("got=test_value_123"));
    }

    [Fact]
    public async Task Execute_RunScript_InvalidCommand_ReturnsError()
    {
        WriteMacro("bad-cmd.yaml", @"
name: Bad Cmd
steps:
  - action: run_script
    command: this_command_does_not_exist_12345
");
        using var engine = new MacroEngine(_tempDir, enableWatcher: false);
        var result = await engine.ExecuteAsync("bad-cmd");
        Assert.False(result.Success);
        Assert.Contains("Failed to start", result.Message);
    }

    [Fact]
    public async Task Execute_RunScript_Timeout_ReturnsError()
    {
        WriteMacro("timeout-test.yaml", @"
name: Timeout Test
steps:
  - action: run_script
    command: cmd.exe
    arguments: /c ping -n 30 127.0.0.1
    timeout: 2
");
        using var engine = new MacroEngine(_tempDir, enableWatcher: false);
        var result = await engine.ExecuteAsync("timeout-test");
        Assert.False(result.Success);
        Assert.Contains("timed out", result.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public async Task Execute_RunScript_DoesNotRequireAttachedProcess()
    {
        // run_script should work even when no process is attached
        WriteMacro("no-attach.yaml", @"
name: No Attach
steps:
  - action: run_script
    command: cmd.exe
    arguments: /c echo works
");
        using var engine = new MacroEngine(_tempDir, enableWatcher: false);
        var result = await engine.ExecuteAsync("no-attach");
        Assert.True(result.Success, result.Message);
    }

    // --- Cancellation ---

    [Fact]
    public async Task Execute_AlreadyCancelledToken_ReturnsImmediately()
    {
        WriteMacro("cancel-test.yaml", @"
name: Cancel Test
steps:
  - action: run_script
    command: cmd.exe
    arguments: /c echo should-not-run
");
        using var engine = new MacroEngine(_tempDir, enableWatcher: false);
        using var cts = new CancellationTokenSource();
        cts.Cancel(); // Already cancelled before execution starts

        var result = await engine.ExecuteAsync("cancel-test", cancellation: cts.Token);
        Assert.False(result.Success);
        Assert.Contains("cancelled", result.Message, StringComparison.OrdinalIgnoreCase);
        Assert.Equal("Cancelled", result.Error);
        Assert.Equal(0, result.StepsExecuted);
    }

    [Fact]
    public async Task Execute_CancelDuringWaitStep_StopsMacro()
    {
        WriteMacro("cancel-wait.yaml", @"
name: Cancel Wait
steps:
  - action: run_script
    command: cmd.exe
    arguments: /c echo step1
  - action: wait
    seconds: 30
  - action: run_script
    command: cmd.exe
    arguments: /c echo should-not-run
");
        using var engine = new MacroEngine(_tempDir, enableWatcher: false);
        using var cts = new CancellationTokenSource();

        // Cancel after 500ms — step 1 completes fast, step 2 (30s wait) gets cancelled
        cts.CancelAfter(TimeSpan.FromMilliseconds(500));

        var result = await engine.ExecuteAsync("cancel-wait", cancellation: cts.Token);
        Assert.False(result.Success);
        Assert.Contains("cancelled", result.Message, StringComparison.OrdinalIgnoreCase);
        Assert.Equal(1, result.StepsExecuted); // Only step 1 completed
    }

    [Fact]
    public async Task Execute_CancelDuringRunScript_KillsProcess()
    {
        // Use a long-running command (ping localhost with count=100)
        WriteMacro("cancel-script.yaml", @"
name: Cancel Script
steps:
  - action: run_script
    command: cmd.exe
    arguments: /c ping -n 100 127.0.0.1
    timeout: 60
");
        using var engine = new MacroEngine(_tempDir, enableWatcher: false);
        using var cts = new CancellationTokenSource();

        // Cancel after 1s — the ping takes much longer
        cts.CancelAfter(TimeSpan.FromSeconds(1));

        var result = await engine.ExecuteAsync("cancel-script", cancellation: cts.Token);
        Assert.False(result.Success);
        Assert.Contains("cancelled", result.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public async Task Execute_CancelledLogMessage_ShowsCancelled()
    {
        WriteMacro("cancel-log.yaml", @"
name: Cancel Log
steps:
  - action: wait
    seconds: 30
");
        using var engine = new MacroEngine(_tempDir, enableWatcher: false);
        using var cts = new CancellationTokenSource();
        var logs = new List<string>();

        cts.CancelAfter(TimeSpan.FromMilliseconds(200));

        var result = await engine.ExecuteAsync("cancel-log", cancellation: cts.Token,
            onLog: msg => logs.Add(msg));
        Assert.False(result.Success);
        Assert.Contains(logs, l => l.Contains("CANCELLED"));
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
