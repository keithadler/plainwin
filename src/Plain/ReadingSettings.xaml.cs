using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;

namespace Plain;

/// <summary>
/// The reading settings. A dyslexic reader put it plainly: these are settings, not features, and they cost nothing in
/// minimalism because nobody has to open them. What they buy is somebody able to read the page at all.
/// </summary>
public partial class ReadingSettings : Window
{
    private readonly Settings _settings;

    public ReadingSettings(Settings settings)
    {
        _settings = settings;
        InitializeComponent();

        // Every face installed on this PC, so a reader who has a particular one can use it.
        var faces = Fonts.SystemFontFamilies
            .Select(f => f.Source)
            .Where(name => !name.StartsWith("Segoe Fluent", StringComparison.Ordinal))
            .OrderBy(name => name, StringComparer.CurrentCultureIgnoreCase)
            .ToList();
        faces.Insert(0, "The one Plain ships with");
        foreach (var face in faces) FontPicker.Items.Add(face);
        FontPicker.SelectedItem = _settings.ReadingFont.Length == 0
            ? faces[0]
            : faces.FirstOrDefault(f => f == _settings.ReadingFont) ?? faces[0];

        foreach (var (name, hex) in Settings.Papers) PaperPicker.Items.Add(new ComboBoxItem { Content = name, Tag = hex });
        PaperPicker.SelectedItem = PaperPicker.Items.Cast<ComboBoxItem>()
            .FirstOrDefault(i => (string)i.Tag == _settings.PaperColour) ?? PaperPicker.Items[0];

        foreach (var spacing in Settings.Spacings)
            SpacingPicker.Items.Add(new ComboBoxItem { Content = spacing == 1.0 ? "Normal" : $"{spacing:0.##} times", Tag = spacing });
        SpacingPicker.SelectedItem = SpacingPicker.Items.Cast<ComboBoxItem>()
            .FirstOrDefault(i => Math.Abs((double)i.Tag - _settings.LineSpacing) < 0.01) ?? SpacingPicker.Items[0];

        WidthSlider.Value = Math.Clamp(_settings.TextWidth, WidthSlider.Minimum, WidthSlider.Maximum);
        KeepRecovery.IsChecked = _settings.KeepRecovery;
        CheckUpdates.IsChecked = _settings.CheckForUpdates;

        FontPicker.SelectionChanged += (_, _) => ShowSample();
        PaperPicker.SelectionChanged += (_, _) => ShowSample();
        SpacingPicker.SelectionChanged += (_, _) => ShowSample();
        ShowSample();
    }

    private void OnWidthChanged(object sender, RoutedPropertyChangedEventArgs<double> e)
    {
        if (WidthLabel is not null) WidthLabel.Text = $"{WidthSlider.Value:0} points";
    }

    private void ShowSample()
    {
        if (Sample is null) return;
        Sample.FontFamily = FontPicker.SelectedIndex <= 0
            ? new FontFamily("Segoe UI Variable Text, Segoe UI")
            : new FontFamily((string)FontPicker.SelectedItem);
        Sample.LineHeight = Sample.FontSize * ((SpacingPicker.SelectedItem as ComboBoxItem)?.Tag as double? ?? 1.0) * 1.3;

        var paper = (PaperPicker.SelectedItem as ComboBoxItem)?.Tag as string ?? "";
        if (Sample.Parent is Border border)
            border.Background = paper.Length == 0
                ? (Brush)Application.Current.Resources["Surface"]
                : new SolidColorBrush((Color)ColorConverter.ConvertFromString(paper));
        Sample.Foreground = paper.Length == 0
            ? (Brush)Application.Current.Resources["Ink"]
            : Brushes.Black;   // a tinted paper is always a light one, so the ink on it is dark
    }

    private void OnReset(object sender, RoutedEventArgs e)
    {
        FontPicker.SelectedIndex = 0;
        PaperPicker.SelectedIndex = 0;
        SpacingPicker.SelectedIndex = 0;
        WidthSlider.Value = 860;
        KeepRecovery.IsChecked = true;
        CheckUpdates.IsChecked = true;
        ShowSample();
    }

    private void OnSave(object sender, RoutedEventArgs e)
    {
        _settings.ReadingFont = FontPicker.SelectedIndex <= 0 ? "" : (string)FontPicker.SelectedItem;
        _settings.PaperColour = (PaperPicker.SelectedItem as ComboBoxItem)?.Tag as string ?? "";
        _settings.LineSpacing = (SpacingPicker.SelectedItem as ComboBoxItem)?.Tag as double? ?? 1.0;
        _settings.TextWidth = WidthSlider.Value;
        _settings.KeepRecovery = KeepRecovery.IsChecked == true;
        _settings.CheckForUpdates = CheckUpdates.IsChecked == true;
        DialogResult = true;
    }
}
