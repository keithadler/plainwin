using System.IO;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using System.Windows.Media.Effects;
using System.Windows.Media.Imaging;
using Grid = System.Windows.Controls.Grid;

namespace Plain;

/// <summary>
/// Four 1600x900 announcement cards, rendered at 2x from the same fixtures as the screenshots, so the pictures in the
/// post are the real window and not a drawing of it. Run it with:  --screenshot &lt;folder&gt; --announce &lt;file&gt;...
/// </summary>
public static class Promo
{
    private sealed record Card(string File, string Headline, string Sub, string? Shot, bool Notes = false, string? Mono = null);

    private static readonly Card[] Cards =
    {
        new("1-hero", "Word, Excel and PowerPoint.\nEdited, not rewritten.",
            "Open a file, change what you came to change, and hand back everything else exactly as you found it. One exe, no installer, no account.",
            "quarter"),
        new("2-preserved", "Two words changed.\nNothing else did.",
            "The macros, the pivot tables, the chart, the tracked changes: Plain does not draw them, so it writes them back as the same bytes it read.",
            "review", Notes: true),
        new("3-honest", "It shows what it kept.\nIt refuses to touch it.",
            "No ribbon and no layout engine. A rail lists every part Plain cannot draw, and it declines the edit rather than guess at one.",
            "woodland"),
        new("4-free", "Free. Open source. No account.",
            "No cloud, no telemetry, no network code at all. Save a file unchanged and it comes back byte for byte, checked on every build. Built by Keith Adler.",
            null, Mono: "plain roundtrip quarter.xlsx"),
    };

    public static void Render(string dir, IReadOnlyList<string> files, TextWriter o)
    {
        Screenshots.Active = true;
        var promoDir = Path.Combine(dir, "promo");
        Directory.CreateDirectory(promoDir);
        var icon = new BitmapImage(new Uri("pack://application:,,,/Assets/plain-256.png"));

        foreach (var c in Cards)
        {
            var root = new Grid { Width = 1600, Height = 900 };
            root.Background = new LinearGradientBrush(
                new GradientStopCollection
                {
                    new(Color.FromRgb(0xBF, 0x8A, 0x2B), 0),
                    new(Color.FromRgb(0x8E, 0x62, 0x14), 0.58),
                    new(Color.FromRgb(0x46, 0x30, 0x0E), 1),
                },
                new Point(0, 0), new Point(1, 1));

            var brand = new StackPanel
            {
                Orientation = Orientation.Horizontal,
                Margin = new Thickness(72, 56, 0, 0),
                VerticalAlignment = VerticalAlignment.Top,
                HorizontalAlignment = HorizontalAlignment.Left,
            };
            brand.Children.Add(new Image { Source = icon, Width = 56, Height = 56, Margin = new Thickness(0, 0, 18, 0) });
            brand.Children.Add(new TextBlock
            {
                Text = "Plain for Windows",
                FontSize = 30,
                FontWeight = FontWeights.SemiBold,
                Foreground = Brushes.White,
                VerticalAlignment = VerticalAlignment.Center,
                FontFamily = new FontFamily("Segoe UI Variable Display, Segoe UI"),
            });
            root.Children.Add(brand);

            bool hasShot = c.Shot is not null;
            var text = new StackPanel
            {
                Margin = new Thickness(72, 40, hasShot ? 0 : 72, 0),
                VerticalAlignment = VerticalAlignment.Center,
                HorizontalAlignment = hasShot ? HorizontalAlignment.Left : HorizontalAlignment.Center,
                Width = hasShot ? 720 : 1200,
            };
            text.Children.Add(new TextBlock
            {
                Text = c.Headline,
                FontSize = hasShot ? 48 : 62,
                FontWeight = FontWeights.Bold,
                Foreground = Brushes.White,
                TextWrapping = TextWrapping.Wrap,
                LineHeight = hasShot ? 60 : 76,
                FontFamily = new FontFamily("Segoe UI Variable Display, Segoe UI"),
                TextAlignment = hasShot ? TextAlignment.Left : TextAlignment.Center,
            });
            text.Children.Add(new TextBlock
            {
                Text = c.Sub,
                FontSize = 25,
                Foreground = new SolidColorBrush(Color.FromArgb(0xE4, 0xFF, 0xFF, 0xFF)),
                TextWrapping = TextWrapping.Wrap,
                Margin = new Thickness(0, 26, 0, 0),
                LineHeight = 36,
                TextAlignment = hasShot ? TextAlignment.Left : TextAlignment.Center,
            });
            if (c.Mono is not null)
            {
                var pill = new Border
                {
                    Background = new SolidColorBrush(Color.FromArgb(0x3A, 0xFF, 0xFF, 0xFF)),
                    CornerRadius = new CornerRadius(12),
                    Padding = new Thickness(28, 14, 28, 14),
                    Margin = new Thickness(0, 44, 0, 0),
                    HorizontalAlignment = HorizontalAlignment.Center,
                };
                pill.Child = new TextBlock
                {
                    Text = c.Mono,
                    FontSize = 28,
                    Foreground = Brushes.White,
                    FontFamily = new FontFamily("Cascadia Mono, Consolas"),
                };
                text.Children.Add(pill);
            }
            root.Children.Add(text);

            if (c.Shot is not null)
            {
                var path = files.FirstOrDefault(f => Path.GetFileNameWithoutExtension(f).Equals(c.Shot, StringComparison.OrdinalIgnoreCase));
                if (path is null || !File.Exists(path)) { o.WriteLine($"no fixture named {c.Shot} for card {c.File}"); continue; }
                var shot = Shot(path, c.Notes);
                root.Children.Add(new Border
                {
                    Width = 790,
                    Height = 490,
                    CornerRadius = new CornerRadius(14),
                    ClipToBounds = true,
                    HorizontalAlignment = HorizontalAlignment.Right,
                    VerticalAlignment = VerticalAlignment.Center,
                    Margin = new Thickness(0, 60, 52, 0),
                    Background = new ImageBrush(shot) { Stretch = Stretch.Uniform },
                    Effect = new DropShadowEffect { BlurRadius = 60, ShadowDepth = 18, Opacity = 0.5, Direction = 270 },
                });
            }

            root.Children.Add(new TextBlock
            {
                Text = "github.com/keithadler/plainwin",
                FontSize = 22,
                Foreground = new SolidColorBrush(Color.FromArgb(0xB8, 0xFF, 0xFF, 0xFF)),
                Margin = new Thickness(0, 0, 72, 44),
                HorizontalAlignment = HorizontalAlignment.Right,
                VerticalAlignment = VerticalAlignment.Bottom,
                FontFamily = new FontFamily("Cascadia Mono, Consolas"),
            });

            root.Measure(new Size(1600, 900));
            root.Arrange(new Rect(0, 0, 1600, 900));
            root.UpdateLayout();

            var bmp = new RenderTargetBitmap(3200, 1800, 192, 192, PixelFormats.Pbgra32);
            var visual = new DrawingVisual();
            using (var dc = visual.RenderOpen())
                dc.DrawRectangle(new VisualBrush(root) { Stretch = Stretch.None }, null, new Rect(0, 0, 1600, 900));
            bmp.Render(visual);

            var outPath = Path.Combine(promoDir, c.File + ".png");
            var encoder = new PngBitmapEncoder();
            encoder.Frames.Add(BitmapFrame.Create(bmp));
            using (var stream = File.Create(outPath)) encoder.Save(stream);
            o.WriteLine($"wrote {outPath}");
        }
    }

    /// <summary>The real window, laid out off screen and copied through a VisualBrush the way the screenshots are.</summary>
    private static BitmapSource Shot(string file, bool notes)
    {
        const int width = 1420, height = 880;
        var window = new MainWindow
        {
            Width = width,
            Height = height,
            WindowStartupLocation = WindowStartupLocation.Manual,
            Left = -10_000,
            Top = -10_000,
            ShowActivated = false,
        };
        window.OpenForScreenshot(file, notes);
        window.Show();
        var content = (FrameworkElement)window.Content;
        content.Measure(new Size(width, height));
        content.Arrange(new Rect(0, 0, width, height));
        content.UpdateLayout();
        var target = new RenderTargetBitmap(width, height, 96, 96, PixelFormats.Pbgra32);
        var visual = new DrawingVisual();
        using (var dc = visual.RenderOpen())
            dc.DrawRectangle(new VisualBrush(content) { Stretch = Stretch.None }, null, new Rect(0, 0, width, height));
        target.Render(visual);
        window.Close();
        return target;
    }
}
