using System.IO;
using Microsoft.Win32;
using Path = System.IO.Path;

namespace Plain;

/// <summary>
/// Letting Explorer open files with Plain.
///
/// This is the one thing Plain does that changes something outside its own folder, so it is careful about it. It
/// writes only under HKEY_CURRENT_USER, so it never needs an administrator and never touches anyone else's account.
/// It does not make itself the default for anything: Windows decides defaults, and an app that seizes them is the
/// kind of app this one exists as an alternative to. What it adds is an "Edit in Plain" line on the right-click
/// menu for the six file kinds Plain understands, and an entry in "Open with" so Plain is offered there.
///
/// Everything it writes it can take away again, and Undo removes exactly what Register wrote and nothing else.
/// </summary>
public static class Explorer
{
    /// <summary>The kinds Plain can open. Anything else is none of its business.</summary>
    public static readonly string[] Extensions = { ".docx", ".docm", ".xlsx", ".xlsm", ".pptx", ".pptm" };

    private const string Verb = "Edit in Plain";
    private const string ProgId = "Plain.Document";

    /// <summary>
    /// The window's exe, quoted, ready to be a command. Explorer must open the window, so when this is asked from
    /// the console twin it looks for the window beside it rather than registering plain.exe, which would open a
    /// console that closes again.
    /// </summary>
    public static string? WindowExe()
    {
        var running = Environment.ProcessPath;
        if (running is not null && Path.GetFileName(running).Equals("Plain for Windows.exe", StringComparison.OrdinalIgnoreCase))
            return running;

        var beside = Path.Combine(Path.GetDirectoryName(running ?? AppContext.BaseDirectory) ?? ".", "Plain for Windows.exe");
        return File.Exists(beside) ? beside : null;
    }

    private static string Command => $"\"{WindowExe()}\" \"%1\"";

    /// <summary>Is the right-click line there, and pointing at this copy of Plain?</summary>
    public static bool Registered()
    {
        try
        {
            using var key = Registry.CurrentUser.OpenSubKey($@"Software\Classes\{ProgId}\shell\open\command");
            return key?.GetValue("") as string == Command;
        }
        catch { return false; }
    }

    /// <summary>Is it there at all, even pointing at a copy of Plain that has since moved or gone?</summary>
    public static bool RegisteredAnywhere()
    {
        try
        {
            using var key = Registry.CurrentUser.OpenSubKey($@"Software\Classes\{ProgId}");
            return key is not null;
        }
        catch { return false; }
    }

    /// <summary>
    /// Add the right-click line and offer Plain under "Open with". Returns what to tell the person, or throws
    /// nothing: a registry that will not be written to is a message, not a crash.
    /// </summary>
    public static string Register()
    {
        if (WindowExe() is null)
            return "Plain could not find \"Plain for Windows.exe\" to point the menu at. Keep it beside plain.exe.";
        try
        {
            using (var progId = Registry.CurrentUser.CreateSubKey($@"Software\Classes\{ProgId}"))
            {
                progId.SetValue("", "Office file");
                progId.SetValue("FriendlyTypeName", "Office file");
            }
            using (var command = Registry.CurrentUser.CreateSubKey($@"Software\Classes\{ProgId}\shell\open\command"))
                command.SetValue("", Command);

            foreach (var extension in Extensions)
            {
                // The right-click line, on the file kind itself. This adds a line; it does not take the default.
                using (var verb = Registry.CurrentUser.CreateSubKey(
                    $@"Software\Classes\SystemFileAssociations\{extension}\shell\{Verb}\command"))
                    verb.SetValue("", Command);

                // And offer Plain in the "Open with" list, which is a suggestion Windows may or may not show.
                using (var open = Registry.CurrentUser.CreateSubKey(
                    $@"Software\Classes\{extension}\OpenWithProgids"))
                    open.SetValue(ProgId, Array.Empty<byte>(), RegistryValueKind.None);
            }

            Refresh();
            return "Plain is now on the right-click menu for Word, Excel and PowerPoint files, as \"Edit in Plain\". "
                 + "It has not made itself the default for anything.";
        }
        catch (Exception ex)
        {
            return "Could not add Plain to the right-click menu: " + ex.Message;
        }
    }

    /// <summary>Take back exactly what Register wrote, leaving anything else alone.</summary>
    public static string Undo()
    {
        try
        {
            foreach (var extension in Extensions)
            {
                Registry.CurrentUser.DeleteSubKeyTree(
                    $@"Software\Classes\SystemFileAssociations\{extension}\shell\{Verb}", throwOnMissingSubKey: false);

                using var open = Registry.CurrentUser.OpenSubKey($@"Software\Classes\{extension}\OpenWithProgids", writable: true);
                if (open?.GetValue(ProgId) is not null) open.DeleteValue(ProgId, throwOnMissingValue: false);
            }
            Registry.CurrentUser.DeleteSubKeyTree($@"Software\Classes\{ProgId}", throwOnMissingSubKey: false);

            Refresh();
            return "Plain has taken itself off the right-click menu.";
        }
        catch (Exception ex)
        {
            return "Could not take Plain off the right-click menu: " + ex.Message;
        }
    }

    /// <summary>Tell Explorer to look again, so the change shows without signing out.</summary>
    private static void Refresh()
    {
        try { SHChangeNotify(0x08000000, 0x0000, IntPtr.Zero, IntPtr.Zero); } catch { /* cosmetic only */ }
    }

    [System.Runtime.InteropServices.DllImport("shell32.dll")]
    private static extern void SHChangeNotify(int eventId, uint flags, IntPtr item1, IntPtr item2);
}
