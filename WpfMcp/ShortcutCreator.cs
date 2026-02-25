using System.IO;

namespace WpfMcp;

/// <summary>
/// Creates Windows .lnk shortcut files using COM interop (WScript.Shell).
/// Used to export macros as double-clickable shortcuts.
/// </summary>
public static class ShortcutCreator
{
    /// <summary>
    /// Create a Windows shortcut (.lnk) file.
    /// </summary>
    /// <param name="lnkPath">Full path for the .lnk file to create.</param>
    /// <param name="targetExe">Full path to the target executable.</param>
    /// <param name="arguments">Command-line arguments for the target.</param>
    /// <param name="workingDirectory">Working directory for the shortcut.</param>
    /// <param name="description">Shortcut description (tooltip).</param>
    /// <param name="runAsAdmin">Set the "Run as administrator" flag on the shortcut.</param>
    /// <returns>The full path of the created .lnk file.</returns>
    public static string CreateShortcut(
        string lnkPath,
        string targetExe,
        string arguments,
        string workingDirectory,
        string description,
        bool runAsAdmin = false)
    {
        // Ensure output directory exists
        var dir = Path.GetDirectoryName(lnkPath);
        if (!string.IsNullOrEmpty(dir) && !Directory.Exists(dir))
            Directory.CreateDirectory(dir);

        // Use COM interop via WScript.Shell to create the .lnk file
        var shellType = Type.GetTypeFromProgID("WScript.Shell")
            ?? throw new InvalidOperationException("WScript.Shell COM object not available");
        dynamic shell = Activator.CreateInstance(shellType)
            ?? throw new InvalidOperationException("Failed to create WScript.Shell instance");

        try
        {
            dynamic shortcut = shell.CreateShortcut(lnkPath);
            try
            {
                shortcut.TargetPath = targetExe;
                shortcut.Arguments = arguments;
                shortcut.WorkingDirectory = workingDirectory;
                shortcut.Description = description;
                shortcut.Save();
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(shortcut);
            }
        }
        finally
        {
            System.Runtime.InteropServices.Marshal.ReleaseComObject(shell);
        }

        // Set "Run as administrator" flag via binary patch.
        // The .lnk file header has a flags field; byte at offset 0x15 controls
        // the "run as admin" bit (0x20). COM API doesn't expose this property,
        // so we patch it directly after Save().
        if (runAsAdmin && File.Exists(lnkPath))
            SetRunAsAdminFlag(lnkPath);

        return lnkPath;
    }

    /// <summary>
    /// Patch the .lnk file to set the "Run as administrator" flag.
    /// Byte at offset 0x15 has bit 0x20 for the SLDF_RUNAS_USER flag.
    /// </summary>
    private static void SetRunAsAdminFlag(string lnkPath)
    {
        var bytes = File.ReadAllBytes(lnkPath);
        if (bytes.Length > 0x15)
        {
            bytes[0x15] |= 0x20;
            File.WriteAllBytes(lnkPath, bytes);
        }
    }
}
