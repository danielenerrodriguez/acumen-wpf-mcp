using System.IO;
using Xunit;

namespace WpfMcp.Tests;

public class MacroExportTests : IDisposable
{
    private readonly string _tempDir;
    private readonly string _macrosDir;
    private readonly string _shortcutsDir;

    public MacroExportTests()
    {
        _tempDir = Path.Combine(Path.GetTempPath(), $"wpfmcp_export_tests_{Guid.NewGuid():N}");
        _macrosDir = Path.Combine(_tempDir, "macros");
        _shortcutsDir = Path.Combine(_tempDir, "shortcuts");
        Directory.CreateDirectory(_macrosDir);
        Directory.CreateDirectory(_shortcutsDir);
    }

    public void Dispose()
    {
        try { Directory.Delete(_tempDir, recursive: true); } catch { }
    }

    private void WriteMacro(string relativePath, string yaml)
    {
        var fullPath = Path.Combine(_macrosDir, relativePath);
        Directory.CreateDirectory(Path.GetDirectoryName(fullPath)!);
        File.WriteAllText(fullPath, yaml);
    }

    // --- ExportMacro ---

    [Fact]
    public void ExportMacro_NonexistentMacro_ReturnsError()
    {
        using var engine = new MacroEngine(_macrosDir, enableWatcher: false);
        var result = engine.ExportMacro("nonexistent", _shortcutsDir);
        Assert.False(result.Ok);
        Assert.Contains("not found", result.Message);
    }

    [Fact]
    public void ExportMacro_ValidMacro_CreatesLnkFile()
    {
        WriteMacro("test.yaml", @"
name: Test
description: A test macro
steps:
  - action: focus
");

        using var engine = new MacroEngine(_macrosDir, enableWatcher: false);
        var result = engine.ExportMacro("test", _shortcutsDir);

        Assert.True(result.Ok, result.Message);
        Assert.Equal("test", result.MacroName);
        Assert.True(File.Exists(result.ShortcutPath), $"Shortcut not found at {result.ShortcutPath}");
        Assert.EndsWith(".lnk", result.ShortcutPath);
    }

    [Fact]
    public void ExportMacro_SubfolderMacro_CreatesLnkInSubfolder()
    {
        WriteMacro(Path.Combine("product", "workflow.yaml"), @"
name: Workflow
description: A workflow macro
steps:
  - action: focus
");

        using var engine = new MacroEngine(_macrosDir, enableWatcher: false);
        var result = engine.ExportMacro("product/workflow", _shortcutsDir);

        Assert.True(result.Ok, result.Message);
        Assert.Contains("product", result.ShortcutPath);
        Assert.True(File.Exists(result.ShortcutPath));
    }

    [Fact]
    public void ExportMacro_ExistingShortcut_WithoutForce_ReturnsError()
    {
        WriteMacro("test.yaml", @"
name: Test
description: A test macro
steps:
  - action: focus
");

        using var engine = new MacroEngine(_macrosDir, enableWatcher: false);

        // Export first time
        var result1 = engine.ExportMacro("test", _shortcutsDir);
        Assert.True(result1.Ok);

        // Export again without force
        var result2 = engine.ExportMacro("test", _shortcutsDir, force: false);
        Assert.False(result2.Ok);
        Assert.Contains("already exists", result2.Message);
    }

    [Fact]
    public void ExportMacro_ExistingShortcut_WithForce_Overwrites()
    {
        WriteMacro("test.yaml", @"
name: Test
description: A test macro
steps:
  - action: focus
");

        using var engine = new MacroEngine(_macrosDir, enableWatcher: false);

        // Export first time
        var result1 = engine.ExportMacro("test", _shortcutsDir);
        Assert.True(result1.Ok);
        var firstWriteTime = File.GetLastWriteTimeUtc(result1.ShortcutPath);

        // Small delay to ensure different timestamp
        Thread.Sleep(50);

        // Export again with force
        var result2 = engine.ExportMacro("test", _shortcutsDir, force: true);
        Assert.True(result2.Ok);
        var secondWriteTime = File.GetLastWriteTimeUtc(result2.ShortcutPath);
        Assert.True(secondWriteTime >= firstWriteTime);
    }

    [Fact]
    public void ExportMacro_LnkFileHasContent()
    {
        WriteMacro("test.yaml", @"
name: Test
description: A test macro
steps:
  - action: focus
");

        using var engine = new MacroEngine(_macrosDir, enableWatcher: false);
        var result = engine.ExportMacro("test", _shortcutsDir);

        Assert.True(result.Ok);
        var bytes = File.ReadAllBytes(result.ShortcutPath);
        // .lnk files start with a 4-byte header class ID
        // The magic bytes for .lnk are: 4C 00 00 00 (76 in decimal)
        Assert.True(bytes.Length > 0x15, "Shortcut file too small");
        Assert.Equal(0x4C, bytes[0]); // .lnk magic byte
    }

    [Fact]
    public void ExportMacro_NoRunAsAdminFlag()
    {
        WriteMacro("test.yaml", @"
name: Test
description: A test macro
steps:
  - action: focus
");

        using var engine = new MacroEngine(_macrosDir, enableWatcher: false);
        var result = engine.ExportMacro("test", _shortcutsDir);

        Assert.True(result.Ok);
        var bytes = File.ReadAllBytes(result.ShortcutPath);
        // Exported shortcuts should NOT have the run-as-admin flag.
        // The drag-and-drop mode handles elevation internally via EnsureServerAndConnectAsync,
        // only prompting for UAC when the elevated server isn't already running.
        Assert.True((bytes[0x15] & 0x20) == 0, "Run as admin flag should not be set");
    }

    // --- ExportAllMacros ---

    [Fact]
    public void ExportAllMacros_MultipleMacros_ExportsAll()
    {
        WriteMacro("macro1.yaml", @"
name: Macro 1
description: First
steps:
  - action: focus
");
        WriteMacro("macro2.yaml", @"
name: Macro 2
description: Second
steps:
  - action: focus
");
        WriteMacro(Path.Combine("sub", "macro3.yaml"), @"
name: Macro 3
description: Third in subfolder
steps:
  - action: focus
");

        using var engine = new MacroEngine(_macrosDir, enableWatcher: false);
        var results = engine.ExportAllMacros(_shortcutsDir);

        Assert.Equal(3, results.Count);
        Assert.All(results, r => Assert.True(r.Ok, r.Message));
        Assert.All(results, r => Assert.True(File.Exists(r.ShortcutPath)));
    }

    [Fact]
    public void ExportAllMacros_EmptyEngine_ReturnsEmptyList()
    {
        using var engine = new MacroEngine(_macrosDir, enableWatcher: false);
        var results = engine.ExportAllMacros(_shortcutsDir);
        Assert.Empty(results);
    }

    // --- ShortcutCreator ---

    [Fact]
    public void ShortcutCreator_CreateShortcut_CreatesFile()
    {
        var lnkPath = Path.Combine(_shortcutsDir, "test-creator.lnk");
        var result = ShortcutCreator.CreateShortcut(
            lnkPath: lnkPath,
            targetExe: @"C:\Windows\notepad.exe",
            arguments: "test.txt",
            workingDirectory: @"C:\Windows",
            description: "Test shortcut",
            runAsAdmin: false);

        Assert.Equal(lnkPath, result);
        Assert.True(File.Exists(lnkPath));
    }

    [Fact]
    public void ShortcutCreator_CreateShortcut_WithRunAsAdmin_SetsFlag()
    {
        var lnkPath = Path.Combine(_shortcutsDir, "admin-test.lnk");
        ShortcutCreator.CreateShortcut(
            lnkPath: lnkPath,
            targetExe: @"C:\Windows\notepad.exe",
            arguments: "",
            workingDirectory: @"C:\Windows",
            description: "Admin shortcut",
            runAsAdmin: true);

        var bytes = File.ReadAllBytes(lnkPath);
        Assert.True((bytes[0x15] & 0x20) != 0, "Run as admin flag not set");
    }

    [Fact]
    public void ShortcutCreator_CreateShortcut_WithoutRunAsAdmin_NoFlag()
    {
        var lnkPath = Path.Combine(_shortcutsDir, "normal-test.lnk");
        ShortcutCreator.CreateShortcut(
            lnkPath: lnkPath,
            targetExe: @"C:\Windows\notepad.exe",
            arguments: "",
            workingDirectory: @"C:\Windows",
            description: "Normal shortcut",
            runAsAdmin: false);

        var bytes = File.ReadAllBytes(lnkPath);
        Assert.True((bytes[0x15] & 0x20) == 0, "Run as admin flag should not be set");
    }

    [Fact]
    public void ShortcutCreator_CreateShortcut_CreatesSubdirectories()
    {
        var lnkPath = Path.Combine(_shortcutsDir, "deep", "nested", "test.lnk");
        ShortcutCreator.CreateShortcut(
            lnkPath: lnkPath,
            targetExe: @"C:\Windows\notepad.exe",
            arguments: "",
            workingDirectory: @"C:\Windows",
            description: "Nested shortcut",
            runAsAdmin: false);

        Assert.True(File.Exists(lnkPath));
    }

    // --- Constants.ResolveShortcutsPath ---

    [Fact]
    public void ResolveShortcutsPath_ExplicitPath_UsesExplicit()
    {
        var path = Constants.ResolveShortcutsPath(@"C:\MyShortcuts");
        Assert.Equal(@"C:\MyShortcuts", path);
    }

    [Fact]
    public void ResolveShortcutsPath_NullPath_ReturnsDefault()
    {
        var path = Constants.ResolveShortcutsPath(null);
        // Should end with "Shortcuts" (the default folder name)
        Assert.EndsWith("Shortcuts", path);
    }

    [Fact]
    public void ResolveShortcutsPath_WithMacrosPath_IsSibling()
    {
        var path = Constants.ResolveShortcutsPath(null, @"C:\WpfMcp\publish\macros");
        // Shortcuts should be sibling of macros: C:\WpfMcp\publish\Shortcuts
        Assert.Equal(@"C:\WpfMcp\publish\Shortcuts", path);
    }

    [Fact]
    public void ExportMacro_DefaultPath_IsSiblingOfMacros()
    {
        WriteMacro("test.yaml", @"
name: Test
description: A test macro
steps:
  - action: focus
");

        using var engine = new MacroEngine(_macrosDir, enableWatcher: false);
        // Export without explicit shortcuts path — should use default (sibling of macros)
        var result = engine.ExportMacro("test");

        Assert.True(result.Ok, result.Message);
        // The shortcut should be in _tempDir/Shortcuts/ (sibling of _tempDir/macros/)
        var expectedDir = Path.Combine(_tempDir, "Shortcuts");
        Assert.StartsWith(expectedDir, result.ShortcutPath);
        Assert.True(File.Exists(result.ShortcutPath));
    }

    // --- ExportMacroResult record ---

    [Fact]
    public void ExportMacroResult_PropertiesAreAccessible()
    {
        var result = new ExportMacroResult(true, @"C:\Shortcuts\test.lnk", "test", "Success");
        Assert.True(result.Ok);
        Assert.Equal(@"C:\Shortcuts\test.lnk", result.ShortcutPath);
        Assert.Equal("test", result.MacroName);
        Assert.Equal("Success", result.Message);
    }
}
