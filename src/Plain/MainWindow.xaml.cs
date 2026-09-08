using System.IO;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using Microsoft.Win32;
using Plain.Core;
using Path = System.IO.Path;

namespace Plain;

public partial class MainWindow : Window
{
    private readonly List<OpenFile> _open = new();
    private OpenFile? _active;
    private bool _railVisible = true;

    /// <summary>One file the app is holding: the model, the view, what changed, and how to put it back.</summary>
    private sealed class OpenFile
    {
        public required PlainFile File { get; init; }
        public required UIElement View { get; init; }
        public required string FilePath { get; init; }
        public Stack<Action> Undo { get; } = new();
        public bool Dirty { get; set; }
        public bool Flattened { get; set; }
        public string Name => Path.GetFileName(FilePath);
    }

    public MainWindow()
    {
        InitializeComponent();
        Loaded += (_, _) =>
        {
            if (!Screenshots.Active)
                foreach (var arg in Environment.GetCommandLineArgs().Skip(1))
                    if (!arg.StartsWith('-') && File.Exists(arg)) OpenPath(arg);
            Refresh();
        };
        KeyDown += OnWindowKey;
    }

    /// <summary>Open a file straight away, for the screenshot renderer which has no user to click Open.</summary>
    internal void OpenForScreenshot(string path)
    {
        OpenPath(path);
        Refresh();
    }

    // ---------- opening ----------

    private void OnOpen(object sender, RoutedEventArgs e)
    {
        var dialog = new OpenFileDialog
        {
            Title = "Open a Word, Excel or PowerPoint file",
            Filter = "Word, Excel and PowerPoint|*.docx;*.docm;*.xlsx;*.xlsm;*.pptx;*.pptm|"
                   + "Word documents|*.docx;*.docm|Excel workbooks|*.xlsx;*.xlsm|PowerPoint decks|*.pptx;*.pptm|All files|*.*",
            Multiselect = true,
        };
        if (dialog.ShowDialog(this) != true) return;
        foreach (var path in dialog.FileNames) OpenPath(path);
        Refresh();
    }

    private void OnDragOver(object sender, DragEventArgs e)
    {
        e.Effects = e.Data.GetDataPresent(DataFormats.FileDrop) ? DragDropEffects.Copy : DragDropEffects.None;
        e.Handled = true;
    }

    private void OnDrop(object sender, DragEventArgs e)
    {
        if (e.Data.GetData(DataFormats.FileDrop) is string[] paths)
        {
            foreach (var path in paths) OpenPath(path);
            Refresh();
        }
    }

    private void OpenPath(string path)
    {
        var already = _open.FirstOrDefault(f => string.Equals(f.FilePath, path, StringComparison.OrdinalIgnoreCase));
        if (already is not null) { _active = already; return; }

        PlainFile file;
        try { file = PlainFile.Open(path); }
        catch (Exception ex) { Say(Path.GetFileName(path) + ": " + Explain(ex)); return; }

        OpenFile entry;
        try { entry = Build(file, path); }
        catch (Exception ex) { Say(Path.GetFileName(path) + ": " + Explain(ex)); return; }

        _open.Add(entry);
        _active = entry;
    }

    private OpenFile Build(PlainFile file, string path)
    {
        UIElement view;
        OpenFile entry;

        switch (file.Kind)
        {
            case FileKind.Spreadsheet:
            {
                var sheetView = new SheetView(file.Workbook!.Sheets[0]);
                view = sheetView;
                entry = new OpenFile { File = file, View = view, FilePath = path };
                sheetView.SelectionChanged += (reference, cell) =>
                {
                    CellRefText.Text = reference.ToString();
                    CellEditor.Text = cell.Formula ?? cell.Raw;
                };
                sheetView.Edited += undo => { entry.Undo.Push(undo); entry.Dirty = true; Refresh(); };
                break;
            }
            case FileKind.Document:
            {
                var docView = new DocView(file.Document!);
                view = docView;
                entry = new OpenFile { File = file, View = view, FilePath = path };
                docView.Edited += undo =>
                {
                    entry.Undo.Push(undo);
                    entry.Dirty = true;
                    entry.Flattened = docView.FlattenedSomething;
                    Refresh();
                };
                break;
            }
            default:
            {
                var deckView = new DeckView(file.Deck!);
                view = deckView;
                entry = new OpenFile { File = file, View = view, FilePath = path };
                deckView.Edited += undo =>
                {
                    entry.Undo.Push(undo);
                    entry.Dirty = true;
                    entry.Flattened = deckView.FlattenedSomething;
                    Refresh();
                };
                break;
            }
        }
        return entry;
    }

    // ---------- commands ----------

    private void OnWindowKey(object sender, KeyEventArgs e)
    {
        if (e.KeyboardDevice.Modifiers != ModifierKeys.Control) return;
        switch (e.Key)
        {
            case Key.O: OnOpen(sender, e); e.Handled = true; break;
            case Key.S: OnSave(sender, e); e.Handled = true; break;
            case Key.Z: OnUndo(sender, e); e.Handled = true; break;
            case Key.W: CloseActive(); e.Handled = true; break;
        }
    }

    private void OnSave(object sender, RoutedEventArgs e)
    {
        if (_active is null) return;
        try
        {
            var (total, edited, kept) = _active.File.Counts();
            _active.File.Save();
            _active.Dirty = false;
            _active.Undo.Clear();
            Say($"Saved {_active.Name}. {edited} of {total} parts rewritten, {kept} kept byte for byte.");
        }
        catch (Exception ex) { Say("Could not save: " + Explain(ex)); }
        Refresh();
    }

    private void OnUndo(object sender, RoutedEventArgs e)
    {
        if (_active is null || _active.Undo.Count == 0) return;
        _active.Undo.Pop()();
        _active.Dirty = _active.Undo.Count > 0;
        Refresh();
    }

    private void CloseActive()
    {
        if (_active is null) return;
        if (_active.Dirty)
        {
            var answer = MessageBox.Show(this,
                $"{_active.Name} has changes you have not saved. Save them before closing?",
                "Plain", MessageBoxButton.YesNoCancel, MessageBoxImage.Question);
            if (answer == MessageBoxResult.Cancel) return;
            if (answer == MessageBoxResult.Yes) OnSave(this, new RoutedEventArgs());
        }
        _open.Remove(_active);
        _active = _open.LastOrDefault();
        Refresh();
    }

    private void OnToggleRail(object sender, RoutedEventArgs e)
    {
        _railVisible = !_railVisible;
        Refresh();
    }

    // ---------- the formula bar ----------

    private void OnFormulaKey(object sender, KeyEventArgs e)
    {
        if (_active?.View is not SheetView view) return;
        if (e.Key == Key.Enter) { view.Apply(CellEditor.Text); view.Focus(); e.Handled = true; }
        else if (e.Key == Key.Escape)
        {
            var cell = view.Sheet.Read(view.Selected);
            CellEditor.Text = cell.Formula ?? cell.Raw;
            view.Focus();
            e.Handled = true;
        }
    }

    private void OnFormulaCommit(object sender, RoutedEventArgs e)
    {
        if (_active?.View is SheetView view && CellEditor.IsKeyboardFocusWithin == false) view.Apply(CellEditor.Text);
    }

    // ---------- painting the chrome ----------

    private string _message = "";
    private void Say(string message) { _message = message; Refresh(); }

    private void Refresh()
    {
        BuildTabs();

        Stage.Content = _active?.View;
        Welcome.Visibility = _active is null ? Visibility.Visible : Visibility.Collapsed;
        FormulaBar.Visibility = _active?.File.Kind == FileKind.Spreadsheet ? Visibility.Visible : Visibility.Collapsed;
        Rail.Visibility = _railVisible && _active is not null ? Visibility.Visible : Visibility.Collapsed;

        SaveBtn.IsEnabled = _active?.Dirty == true;
        UndoBtn.IsEnabled = _active?.Undo.Count > 0;
        SaveBtn.Content = _active?.Dirty == true ? "Save" : "Saved";

        ContextHint.Text = _active?.File.Kind switch
        {
            FileKind.Spreadsheet => "Click a cell and type. A leading = makes a formula.",
            FileKind.Document => "Click any line and type. One column, no page breaks.",
            FileKind.Presentation => "Pick a slide, then edit its text.",
            _ => "",
        };

        if (_active is null)
        {
            RailList.ItemsSource = null;
            RailFoot.Text = "";
            KeepBtn.Content = "Preserved";
            StatusPromise.Text = _message.Length > 0 ? _message : "Nothing open.";
            StatusStat.Text = "";
            Title = "Plain";
            return;
        }

        var parts = _active.File.Parts();
        var kept = parts.Where(p => p.Role == PartRole.Preserved).ToList();
        RailList.ItemsSource = kept;
        KeepBtn.Content = $"Preserved  {kept.Count}";
        RailFoot.Text = _active.File.Kind switch
        {
            FileKind.Spreadsheet => "Editing a cell clears the values Excel cached beside formulas, so no total on screen is out of date.",
            FileKind.Document => "Tracked changes, comments, headers and footers stay in the file. Plain shows the text of the body.",
            _ => "Shapes Plain cannot draw are held in place. Editing a title never moves them.",
        };

        int shown = parts.Count(p => p.Role == PartRole.Shown);
        StatusPromise.Text = _message.Length > 0
            ? _message
            : $"{parts.Count} parts read, {shown} shown, {parts.Count - shown} kept byte for byte";

        StatusStat.Text = _active.File.Kind switch
        {
            FileKind.Spreadsheet => Describe(_active),
            FileKind.Document => $"{_active.File.Document!.BlockCount} blocks, flowing view with no page breaks",
            _ => $"Slide {(_active.View as DeckView)?.Current.Number} of {_active.File.Deck!.Slides.Count}",
        };

        if (_active.Flattened)
            StatusStat.Text += "  ·  one block's mixed formatting was flattened";

        Title = (_active.Dirty ? "• " : "") + _active.Name + " - Plain";
    }

    private static string Describe(OpenFile file)
    {
        var sheet = (file.View as SheetView)?.Sheet;
        if (sheet is null) return "";
        var extent = sheet.Extent;
        return $"{sheet.Name}, used to {extent}";
    }

    private void BuildTabs()
    {
        var strip = new List<UIElement>();
        foreach (var file in _open)
        {
            bool on = file == _active;
            var label = new TextBlock
            {
                Text = file.Name,
                MaxWidth = 210,
                TextTrimming = TextTrimming.CharacterEllipsis,
                FontWeight = on ? FontWeights.SemiBold : FontWeights.Normal,
                Foreground = on ? App.B("Ink") : App.B("Ink2"),
                VerticalAlignment = VerticalAlignment.Center,
            };
            var kind = new Border
            {
                Background = on ? App.B("AccentSoft") : App.B("Surface2"),
                CornerRadius = new CornerRadius(3),
                Padding = new Thickness(4, 2, 4, 2),
                Margin = new Thickness(0, 0, 8, 0),
                VerticalAlignment = VerticalAlignment.Center,
                Child = new TextBlock
                {
                    Text = Path.GetExtension(file.FilePath).TrimStart('.').ToUpperInvariant(),
                    FontFamily = new FontFamily("Cascadia Mono, Consolas"),
                    FontSize = 9.5,
                    Foreground = on ? App.B("AccentInk") : App.B("Ink3"),
                },
            };
            var row = new StackPanel { Orientation = Orientation.Horizontal };
            row.Children.Add(kind);
            row.Children.Add(label);
            if (file.Dirty)
                row.Children.Add(new System.Windows.Shapes.Ellipse
                {
                    Width = 6, Height = 6, Fill = App.B("Accent"),
                    Margin = new Thickness(8, 0, 0, 0), VerticalAlignment = VerticalAlignment.Center,
                });

            var tab = new Button
            {
                Content = row,
                Padding = new Thickness(12, 0, 12, 0),
                Height = 33,
                Background = on ? App.B("Surface") : Brushes.Transparent,
                BorderBrush = on ? App.B("Line") : Brushes.Transparent,
                BorderThickness = new Thickness(1, 1, 1, 0),
                Cursor = Cursors.Hand,
                Tag = file,
            };
            tab.Click += (s, _) => { _active = (OpenFile)((Button)s).Tag; _message = ""; Refresh(); };
            tab.MouseRightButtonUp += (s, _) => { _active = (OpenFile)((Button)s).Tag; CloseActive(); };
            strip.Add(tab);
        }
        Tabs.ItemsSource = strip;
    }

    private static string Explain(Exception ex) => ex switch
    {
        OpcPackage.PackageException p => p.Message,
        UnauthorizedAccessException => "Windows would not let Plain open that file.",
        IOException io when io.Message.Contains("being used") => "Another program has that file open.",
        IOException io => io.Message,
        _ => ex.Message,
    };

    protected override void OnClosing(System.ComponentModel.CancelEventArgs e)
    {
        foreach (var file in _open.Where(f => f.Dirty).ToList())
        {
            var answer = MessageBox.Show(this,
                $"{file.Name} has changes you have not saved. Save them before closing?",
                "Plain", MessageBoxButton.YesNoCancel, MessageBoxImage.Question);
            if (answer == MessageBoxResult.Cancel) { e.Cancel = true; return; }
            if (answer == MessageBoxResult.Yes)
            {
                try { file.File.Save(); }
                catch (Exception ex)
                {
                    MessageBox.Show(this, "Could not save " + file.Name + ": " + Explain(ex), "Plain");
                    e.Cancel = true;
                    return;
                }
            }
        }
        base.OnClosing(e);
    }
}
