#if CLI_BUILD
namespace Plain;

/// <summary>
/// The console twin's entry point. The window build has its own, generated from App.xaml, so one source tree gives
/// two programs: "Plain for Windows.exe" to look at a file and "plain.exe" to script one.
/// </summary>
public static class Startup
{
    [STAThread]
    public static int Main(string[] args) => Cli.Run(args, Console.Out, Console.Error);
}
#endif
