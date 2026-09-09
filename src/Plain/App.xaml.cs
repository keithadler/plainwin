using System.Windows;
using System.Windows.Interop;
using System.Windows.Media;
using Microsoft.Win32;

namespace Plain;

public partial class App : Application
{

    protected override void OnStartup(StartupEventArgs e)
    {
        // Some virtual machines and remote desktops cannot present WPF's hardware surface, and the window comes up
        // blank. Setting PLAIN_SOFTWARE_RENDER=1 draws everything on the processor instead, which is slower but shows.
        if (Environment.GetEnvironmentVariable("PLAIN_SOFTWARE_RENDER") == "1")
            RenderOptions.ProcessRenderMode = RenderMode.SoftwareOnly;

        var args = Environment.GetCommandLineArgs();
        int flag = Array.FindIndex(args, a => a == "--screenshot");
        if (flag >= 0 && flag + 2 < args.Length)
        {
            // Render pictures of the window and quit, without ever putting a window on screen.
            ApplyTheme(args.Contains("--dark"));
            Screenshots.ShowNotes = args.Contains("--notes");
            var files = args.Skip(flag + 2).Where(a => !a.StartsWith('-')).ToList();
            Shutdown(Screenshots.Render(args[flag + 1], files, args.Contains("--dark")));
            return;
        }

        base.OnStartup(e);
        ApplyTheme(WindowsPrefersDark());
        new MainWindow().Show();
    }

    /// <summary>Windows keeps the app theme in the registry; Plain follows it rather than adding a setting for it.</summary>
    public static bool WindowsPrefersDark()
    {
        try
        {
            using var key = Registry.CurrentUser.OpenSubKey(@"Software\Microsoft\Windows\CurrentVersion\Themes\Personalize");
            return key?.GetValue("AppsUseLightTheme") is int light && light == 0;
        }
        catch { return false; }
    }

    public static void ApplyTheme(bool dark)
    {
        var set = dark
            ? new (string Key, string Hex)[]
            {
                ("Chrome", "#1E1E1B"), ("Chrome2", "#262521"), ("Surface", "#26251F"), ("Surface2", "#2E2D27"),
                ("Line", "#3A3932"), ("LineStrong", "#4C4A41"), ("Ink", "#EDEBE3"), ("Ink2", "#A7A49A"),
                ("Ink3", "#7B7970"), ("Accent", "#6FB0E8"), ("AccentInk", "#9BCBF3"), ("AccentSoft", "#1A2E3F"),
                ("Preserved", "#E0AE52"), ("PreservedSoft", "#332B1B"), ("PreservedLine", "#574A2E"),
            }
            : new (string Key, string Hex)[]
            {
                ("Chrome", "#F6F5F2"), ("Chrome2", "#EAE8E2"), ("Surface", "#FDFDFB"), ("Surface2", "#F2F1ED"),
                ("Line", "#DEDCD5"), ("LineStrong", "#C7C4BA"), ("Ink", "#1B1A17"), ("Ink2", "#5E5C53"),
                ("Ink3", "#8D8A7F"), ("Accent", "#1F6FB2"), ("AccentInk", "#145084"), ("AccentSoft", "#E2EDF7"),
                ("Preserved", "#9A6B15"), ("PreservedSoft", "#F7EFDC"), ("PreservedLine", "#E3D3AE"),
            };

        foreach (var (key, hex) in set)
            Current.Resources[key] = new SolidColorBrush((Color)ColorConverter.ConvertFromString(hex));
        Current.Resources["IsDark"] = dark;
    }

    public static Brush B(string key) => (Brush)Current.Resources[key];
    public static Color C(string key) => ((SolidColorBrush)Current.Resources[key]).Color;
}
