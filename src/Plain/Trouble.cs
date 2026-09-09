using System.IO;
using System.Windows;
using System.Windows.Threading;

namespace Plain;

/// <summary>
/// What happens when something goes wrong that nobody thought of.
///
/// Without this, an unexpected error closes the window with the .NET crash box, which says nothing a person can
/// act on and takes whatever was unsaved with it. That is the worst possible ending for a program whose whole
/// promise is that it does not damage your work.
///
/// So: keep a copy of anything unsaved first, write what happened to a file beside the settings, then say plainly
/// what happened, where the note is, and that the work was kept. Nothing is sent anywhere. The note is on your
/// machine, and it is yours to read or delete or attach to a bug report.
/// </summary>
public static class Trouble
{
    /// <summary>Where the note goes. Beside the settings, so everything Plain writes is in one place.</summary>
    public static string LogPath => Path.Combine(Settings.Folder, "trouble.txt");

    /// <summary>Asked for the files that are open, so a crash can keep them before it says anything.</summary>
    public static Func<int>? KeepUnsaved { get; set; }

    public static void Watch(Application app)
    {
        app.DispatcherUnhandledException += (_, e) =>
        {
            e.Handled = Handle(e.Exception, "the window");
        };

        // A background thread that throws takes the process down without the dispatcher ever hearing about it,
        // so the note is written from here even though nothing can be recovered afterwards.
        AppDomain.CurrentDomain.UnhandledException += (_, e) =>
        {
            if (e.ExceptionObject is Exception ex) Write(ex, "a background job");
        };

        TaskScheduler.UnobservedTaskException += (_, e) =>
        {
            Write(e.Exception, "a background job nobody was waiting on");
            e.SetObserved();
        };
    }

    /// <summary>
    /// Keep the work, write the note, tell the person. Returns true when the window can carry on, which it usually
    /// can: one operation failing is not a reason to throw away everything else that is open.
    /// </summary>
    private static bool Handle(Exception ex, string where)
    {
        int kept = 0;
        try { kept = KeepUnsaved?.Invoke() ?? 0; } catch { /* the keeping failed too; carry on and say so */ }

        var note = Write(ex, where);

        var said = "Something went wrong inside Plain that it did not expect.\n\n"
                 + Summarise(ex) + "\n\n"
                 + (kept > 0
                     ? $"A copy of {kept} unsaved file{(kept == 1 ? "" : "s")} was kept, and Plain will offer "
                       + "it back the next time it starts.\n\n"
                     : "Nothing was open with unsaved changes.\n\n")
                 + (note is null
                     ? "Plain could not write a note about it.\n\n"
                     : $"What happened is written down in:\n{note}\n\nNothing has been sent anywhere.\n\n")
                 + "Your file on disk has not been changed. You can carry on, but saving anything you were part "
                 + "way through is worth doing now.";

        try
        {
            MessageBox.Show(said, "Plain", MessageBoxButton.OK, MessageBoxImage.Warning);
        }
        catch { return false; }
        return true;
    }

    /// <summary>The note, appended so an error that happens twice shows as twice. Returns where it went.</summary>
    private static string? Write(Exception ex, string where)
    {
        try
        {
            Directory.CreateDirectory(Settings.Folder);
            var text = $"""

                ---- {DateTimeOffset.Now:yyyy-MM-dd HH:mm:ss zzz} ----
                Plain for Windows {Cli.Version}, {Environment.OSVersion.VersionString}, {System.Runtime.InteropServices.RuntimeInformation.OSArchitecture}
                In: {where}
                {ex}

                """;
            File.AppendAllText(LogPath, text);
            return LogPath;
        }
        catch { return null; }
    }

    /// <summary>One line a person can act on, rather than a type name and a stack.</summary>
    private static string Summarise(Exception ex) => ex switch
    {
        UnauthorizedAccessException => "Windows would not let Plain read or write a file it needed.",
        IOException io when io.Message.Contains("being used", StringComparison.OrdinalIgnoreCase)
            => "A file Plain needed is open in another program.",
        IOException => "Something went wrong reading or writing a file.",
        OutOfMemoryException => "The file was too large for the memory available.",
        _ => ex.Message.Length > 0 && ex.Message.Length < 200 ? ex.Message : ex.GetType().Name,
    };
}
