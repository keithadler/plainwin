using System.IO;
using System.Windows;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Threading;

namespace Plain;

/// <summary>
/// Renders the window to a PNG inside the process, without asking the display for anything. That makes the pictures
/// in the README reproducible on any machine, and it works where a virtual machine cannot composite a screenshot.
/// Run it with:  "Plain for Windows.exe" --screenshot &lt;folder&gt; &lt;file&gt;...
/// </summary>
public static class Screenshots
{
    /// <summary>True while rendering, so a window does not also open whatever is on the command line.</summary>
    public static bool Active { get; internal set; }

    /// <summary>Open the comments panel in the picture, so the rendered shots can show it.</summary>
    public static bool ShowNotes { get; set; }

    public static int Render(string outDir, IReadOnlyList<string> files, bool dark)
    {
        Active = true;
        Directory.CreateDirectory(outDir);
        App.ApplyTheme(dark);
        // Closing each window must not end the run; the caller decides when the app exits.
        Application.Current.ShutdownMode = ShutdownMode.OnExplicitShutdown;

        const int width = 1180, height = 760;
        int written = 0;

        foreach (var file in files)
        {
            if (!File.Exists(file)) { Console.Error.WriteLine($"no file at {file}"); continue; }

            var window = new MainWindow
            {
                Width = width,
                Height = height,
                WindowStartupLocation = WindowStartupLocation.Manual,
                Left = -10_000,
                Top = -10_000,
                ShowActivated = false,
            };
            window.OpenForScreenshot(file, ShowNotes);
            window.Show();
            Pump(window.Dispatcher);

            // Measure and arrange explicitly so the tree is laid out even though nothing is on screen.
            var content = (FrameworkElement)window.Content;
            content.Measure(new Size(width, height));
            content.Arrange(new Rect(0, 0, width, height));
            content.UpdateLayout();
            Pump(window.Dispatcher);

            // Drawing through a VisualBrush copies the laid-out tree rather than the window's own render surface,
            // which is what makes this work where a plain screen capture comes back blank.
            var target = new RenderTargetBitmap(width, height, 96, 96, PixelFormats.Pbgra32);
            var visual = new DrawingVisual();
            using (var dc = visual.RenderOpen())
                dc.DrawRectangle(new VisualBrush(content) { Stretch = Stretch.None }, null, new Rect(0, 0, width, height));
            target.Render(visual);

            var name = Path.GetFileNameWithoutExtension(file) + (dark ? "-dark" : "") + ".png";
            var path = Path.Combine(outDir, name);
            var encoder = new PngBitmapEncoder();
            encoder.Frames.Add(BitmapFrame.Create(target));
            using (var stream = File.Create(path)) encoder.Save(stream);

            window.Close();
            Pump(window.Dispatcher);
            Console.WriteLine($"wrote {path}");
            written++;
        }

        return written > 0 ? 0 : 2;
    }

    /// <summary>Let the dispatcher finish the work the window queued, the way a real message loop would.</summary>
    private static void Pump(Dispatcher dispatcher)
    {
        for (int i = 0; i < 3; i++)
        {
            dispatcher.Invoke(() => { }, DispatcherPriority.ContextIdle);
            dispatcher.Invoke(() => { }, DispatcherPriority.Loaded);
            dispatcher.Invoke(() => { }, DispatcherPriority.Render);
        }
    }
}
