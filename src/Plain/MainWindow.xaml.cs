using System.IO;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using Microsoft.Win32;
using Plain.Core;
// Plain.Core has its own Grid, for rows and columns of a sheet; in this file Grid means the WPF panel.
using Grid = System.Windows.Controls.Grid;
using Path = System.IO.Path;

namespace Plain;

public partial class MainWindow : Window
{
    private readonly List<OpenFile> _open = new();
    private readonly Settings _settings = Settings.Load();
    private OpenFile? _active;
    private bool _railVisible = true;
    private bool _notesVisible;
    private bool _picturesVisible;
    private System.Windows.Threading.DispatcherTimer? _keeper;

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
        _railVisible = _settings.ShowPreserved;
        ApplyTextScale();
        StartKeeping();
        Loaded += (_, _) =>
        {
            if (!Screenshots.Active)
            {
                foreach (var arg in Environment.GetCommandLineArgs().Skip(1))
                    if (!arg.StartsWith('-') && File.Exists(arg)) OpenPath(arg);
                OfferRecovery();
            }
            Refresh();
        };
        KeyDown += OnWindowKey;
    }

    /// <summary>Open a file straight away, for the screenshot renderer which has no user to click Open.</summary>
    internal void OpenForScreenshot(string path, bool showNotes = false)
    {
        OpenPath(path);
        _notesVisible = showNotes;
        Refresh();
    }

    // ---------- keeping what has not been saved ----------

    /// <summary>
    /// Every half minute, put a copy of anything unsaved beside the settings. The machines this runs on lose power
    /// without warning, and losing an afternoon of typing to that is the difference between a tool people trust and
    /// one they do not.
    /// </summary>
    private void StartKeeping()
    {
        _keeper = new System.Windows.Threading.DispatcherTimer { Interval = TimeSpan.FromSeconds(30) };
        _keeper.Tick += (_, _) =>
        {
            if (!_settings.KeepRecovery) return;
            foreach (var file in _open.Where(f => f.Dirty))
                Recovery.Keep(file.File, file.FilePath);
        };
        _keeper.Start();
    }

    /// <summary>Offer back anything a previous run did not get to save.</summary>
    private void OfferRecovery()
    {
        var waiting = Recovery.Waiting().Where(w => !_open.Any(f => string.Equals(f.FilePath, w.Original, StringComparison.OrdinalIgnoreCase))).ToList();
        if (waiting.Count == 0) return;

        var names = string.Join("\n", waiting.Select(w => $"  {Path.GetFileName(w.Original)}  (kept {w.When.Replace("T", " at ")})"));
        var answer = MessageBox.Show(this,
            $"Plain closed with changes it had not saved:\n\n{names}\n\nOpen the kept copies? " +
            "They open beside your files, so nothing is written over until you save.",
            "Plain", MessageBoxButton.YesNo, MessageBoxImage.Information);

        if (answer != MessageBoxResult.Yes) { Recovery.ForgetAll(); return; }

        foreach (var (original, _, copy) in waiting)
        {
            try
            {
                var file = PlainFile.Open(copy);
                var entry = Build(file, original);
                entry.Dirty = true;
                _open.Add(entry);
                _active = entry;
            }
            catch (Exception ex) { Say($"Could not open the kept copy of {Path.GetFileName(original)}: {Explain(ex)}"); }
        }
        Recovery.ForgetAll();
        Say("These are the kept copies. Save them to put the changes back into your files.");
    }

    // ---------- opening ----------

    /// <summary>
    /// Make a new file. Which kind comes from the name you give it, so the one dialog picks both the place and the
    /// kind, and there is no menu to walk through first. The file exists on disk before you type into it, which
    /// means there is nothing to lose if the machine gives up half way through your first paragraph.
    /// </summary>
    private void OnNew(object sender, RoutedEventArgs e)
    {
        var dialog = new SaveFileDialog
        {
            Title = "New file",
            FileName = "Untitled.xlsx",
            DefaultExt = ".xlsx",
            Filter = "Excel workbook|*.xlsx|Word document|*.docx|PowerPoint deck|*.pptx",
            OverwritePrompt = false,   // Plain refuses to write over one itself, with a clearer message
        };
        if (dialog.ShowDialog(this) != true) return;

        try
        {
            var made = PlainFile.Create(dialog.FileName);
            var entry = Build(made, dialog.FileName);
            _open.Add(entry);
            _active = entry;
            Say($"Made {entry.Name}. It is on disk already, so there is nothing to lose.");
        }
        catch (Exception ex) { Say("Could not make that file: " + Explain(ex)); }
        Refresh();
    }

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
        _settings.Remember(path);
    }

    private OpenFile Build(PlainFile file, string path)
    {
        UIElement view;
        OpenFile entry;

        switch (file.Kind)
        {
            case FileKind.Spreadsheet:
            {
                var bookView = new WorkbookView(file.Workbook!);
                view = bookView;
                entry = new OpenFile { File = file, View = view, FilePath = path };
                bookView.SelectionChanged += (reference, cell) =>
                {
                    CellRefText.Text = reference.ToString();
                    CellEditor.Text = cell.Formula ?? cell.Raw;
                    if (_active == entry) StatusStat.Text = Describe(entry);
                };
                bookView.Edited += undo => { entry.Undo.Push(undo); entry.Dirty = true; Refresh(); };
                bookView.GridChangeRequested += (edit, at) => ChangeGrid(entry, bookView, edit, at);
                break;
            }
            case FileKind.Document:
            {
                var docView = new DocView(file.Document!);
                docView.ApplyReading(_settings);
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

    /// <summary>One line of the recent list.</summary>
    private sealed record RecentEntry(string Name, string Path);

    private void OnOpenRecent(object sender, RoutedEventArgs e)
    {
        if (sender is Button { Tag: string path }) { OpenPath(path); Refresh(); }
    }

    // ---------- how big the text is ----------

    /// <summary>
    /// Scale everything in the window. Windows' own display scaling handles the whole screen; this is for the person
    /// who wants this app's text larger than the rest, which was the first thing asked for by anyone reading it at a
    /// distance.
    /// </summary>
    private void ApplyTextScale()
    {
        double scale = Math.Clamp(_settings.TextScale, 0.8, 3.0);
        LayoutTransform = scale == 1.0 ? System.Windows.Media.Transform.Identity
                                       : new System.Windows.Media.ScaleTransform(scale, scale);
    }

    private void Bigger() { _settings.Bigger(); _settings.Save(); ApplyTextScale(); Say($"Text at {_settings.TextScale * 100:0}%."); }
    private void Smaller() { _settings.Smaller(); _settings.Save(); ApplyTextScale(); Say($"Text at {_settings.TextScale * 100:0}%."); }
    private void NormalSize() { _settings.TextScale = 1.0; _settings.Save(); ApplyTextScale(); Say("Text back to normal size."); }

    // ---------- moving about without a mouse ----------

    /// <summary>
    /// Move focus from one part of the window to the next. Without this the grid is a trap: Tab inside it moves the
    /// selected cell, so there is no way out with the keyboard alone.
    /// </summary>
    private void CycleFocus(bool backwards)
    {
        var stops = new List<System.Windows.IInputElement?> { NewBtn, Stage.Content as System.Windows.IInputElement, RailList, Tabs };
        if (_notesVisible) stops.Insert(3, NotesList);
        var live = stops.Where(x => x is UIElement { IsVisible: true }).Cast<System.Windows.IInputElement>().ToList();
        if (live.Count == 0) return;

        int at = live.FindIndex(x => x is DependencyObject d && IsAncestorOfFocus(d));
        int next = at < 0 ? 0 : (at + (backwards ? -1 : 1) + live.Count) % live.Count;
        var target = live[next];
        if (target is UIElement element)
        {
            element.Focus();
            if (!element.IsKeyboardFocusWithin) element.MoveFocus(new TraversalRequest(FocusNavigationDirection.First));
        }
        Say(FocusName(live[next]));
    }

    private static bool IsAncestorOfFocus(DependencyObject candidate)
    {
        var focused = Keyboard.FocusedElement as DependencyObject;
        while (focused is not null)
        {
            if (ReferenceEquals(focused, candidate)) return true;
            focused = System.Windows.Media.VisualTreeHelper.GetParent(focused)
                      ?? LogicalTreeHelper.GetParent(focused);
        }
        return false;
    }

    private string FocusName(object element) => element switch
    {
        _ when ReferenceEquals(element, NewBtn) => "Commands",
        _ when ReferenceEquals(element, RailList) => "Preserved panel",
        _ when ReferenceEquals(element, NotesList) => "Comments panel",
        _ when ReferenceEquals(element, Tabs) => "Open files",
        _ => "Document",
    };

    // ---------- commands ----------

    private void OnWindowKey(object sender, KeyEventArgs e)
    {
        if (e.Key == Key.F6)
        {
            CycleFocus(backwards: e.KeyboardDevice.Modifiers == ModifierKeys.Shift);
            e.Handled = true;
            return;
        }
        if (e.KeyboardDevice.Modifiers != ModifierKeys.Control) return;
        switch (e.Key)
        {
            case Key.O: OnOpen(sender, e); e.Handled = true; break;
            case Key.S: OnSave(sender, e); e.Handled = true; break;
            case Key.Z: OnUndo(sender, e); e.Handled = true; break;
            case Key.W: CloseActive(); e.Handled = true; break;
            case Key.F: ShowFind(); e.Handled = true; break;
            case Key.H: ShowFind(replacing: true); e.Handled = true; break;
            case Key.N: OnNew(sender, e); e.Handled = true; break;
            case Key.P: OnPrint(sender, e); e.Handled = true; break;
            case Key.D: FillDown(); e.Handled = true; break;
            case Key.B: Mark("b"); e.Handled = true; break;
            case Key.I: Mark("i"); e.Handled = true; break;
            case Key.OemPlus or Key.Add: Bigger(); e.Handled = true; break;
            case Key.OemMinus or Key.Subtract: Smaller(); e.Handled = true; break;
            case Key.D0 or Key.NumPad0: NormalSize(); e.Handled = true; break;
        }
    }

    private void OnSave(object sender, RoutedEventArgs e)
    {
        if (_active is null) return;

        if (_active.File.IsReadOnly())
        {
            MessageBox.Show(this,
                $"{_active.Name} is marked read only, so Plain cannot write to it. \"Save a copy\" will write your changes somewhere else.",
                "Plain", MessageBoxButton.OK, MessageBoxImage.Information);
            return;
        }

        if (_active.File.ChangedOnDisk())
        {
            var answer = MessageBox.Show(this,
                $"Something else has changed {_active.Name} since you opened it. Saving now would throw those changes away.\n\nSave anyway?",
                "Plain", MessageBoxButton.YesNo, MessageBoxImage.Warning);
            if (answer != MessageBoxResult.Yes) return;
        }

        try
        {
            var (total, edited, kept) = _active.File.Counts();
            _active.File.Save();
            _active.Dirty = false;
            Recovery.Forget(_active.FilePath);
            // The history is kept: saving is not a wall you cannot step back over, and undoing after a save
            // leaves the file differing from what is on disk again, which is what Dirty then says.
            Say($"Saved {_active.Name}. {edited} of {total} parts rewritten, {kept} kept byte for byte.");
        }
        catch (Exception ex) { Say("Could not save: " + Explain(ex)); }
        Refresh();
    }

    /// <summary>
    /// Write the file, with the changes, to a new name and leave the original alone. The copy is a whole file, not a
    /// patch: every part Plain preserved is in it, byte for byte, exactly as in the original.
    /// </summary>
    private void OnSaveCopy(object sender, RoutedEventArgs e)
    {
        if (_active is null) return;
        var extension = Path.GetExtension(_active.FilePath);
        var dialog = new SaveFileDialog
        {
            Title = "Save a copy",
            FileName = Path.GetFileNameWithoutExtension(_active.FilePath) + " copy" + extension,
            DefaultExt = extension,
            Filter = $"Same kind of file (*{extension})|*{extension}|All files|*.*",
            InitialDirectory = Path.GetDirectoryName(_active.FilePath),
            OverwritePrompt = true,
        };
        if (dialog.ShowDialog(this) != true) return;
        try
        {
            var (total, edited, kept) = _active.File.Counts();
            _active.File.Save(dialog.FileName);
            Say($"Wrote {Path.GetFileName(dialog.FileName)}. {edited} of {total} parts rewritten, {kept} copied byte for byte. {_active.Name} is untouched.");
        }
        catch (Exception ex) { Say("Could not write the copy: " + Explain(ex)); }
        Refresh();
    }

    private void OnUndo(object sender, RoutedEventArgs e)
    {
        if (_active is null || _active.Undo.Count == 0) return;
        _active.Undo.Pop()();
        _active.Dirty = true;    // whatever is on disk, the file in front of you has just changed again
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
        _settings.ShowPreserved = _railVisible;
        _settings.Save();
        Refresh();
    }

    private void OnToggleNotes(object sender, RoutedEventArgs e)
    {
        _notesVisible = !_notesVisible;
        if (_notesVisible) _picturesVisible = false;
        Refresh();
    }

    private void OnTogglePictures(object sender, RoutedEventArgs e)
    {
        _picturesVisible = !_picturesVisible;
        if (_picturesVisible) _notesVisible = false;
        Refresh();
    }

    private void OnSavePictures(object sender, RoutedEventArgs e)
    {
        if (_active is null) return;
        var pictures = Media.In(_active.File);
        if (pictures.Count == 0) return;
        var dialog = new SaveFileDialog
        {
            Title = "Choose a folder: the pictures go beside this name",
            FileName = "pictures here.txt",
            InitialDirectory = Path.GetDirectoryName(_active.FilePath),
        };
        if (dialog.ShowDialog(this) != true) return;
        var folder = Path.GetDirectoryName(dialog.FileName)!;
        try
        {
            foreach (var picture in pictures) Media.SaveTo(picture, folder);
            Say($"Wrote {pictures.Count} picture{(pictures.Count == 1 ? "" : "s")} into {folder}.");
        }
        catch (Exception ex) { Say("Could not write the pictures: " + Explain(ex)); }
    }

    /// <summary>Accept one tracked change, or keep one comment's thread by doing nothing to it.</summary>
    private void OnSettleOne(object sender, RoutedEventArgs e) => Settle(sender, keep: true);
    private void OnDropOne(object sender, RoutedEventArgs e) => Settle(sender, keep: false);

    private void Settle(object sender, bool keep)
    {
        if (_active is null || sender is not Button { Tag: string id }) return;
        var parts = id.Split(':');
        if (parts.Length != 2 || !int.TryParse(parts[1], out var index)) return;

        try
        {
            var before = Snapshot(_active);
            bool done = parts[0] switch
            {
                "change" => Annotations.SettleRevision(_active.File, index, accept: keep),
                "comment" => !keep && Annotations.RemoveComment(_active.File, index),
                _ => false,
            };
            if (!done) { if (parts[0] == "comment" && keep) Say("Comments are kept unless you remove them."); return; }
            _active.Dirty = true;
            _active.Undo.Push(() => Restore(_active, before));
            Rebuild(_active);
            Say(parts[0] == "comment" ? "Comment removed." : keep ? "Change accepted." : "Change turned down.");
        }
        catch (Exception ex) { Say("Could not do that: " + Explain(ex)); }
    }

    private void OnRejectChanges(object sender, RoutedEventArgs e)
    {
        if (_active is null) return;
        var before = Snapshot(_active);
        int settled = Annotations.RejectRevisions(_active.File);
        if (settled == 0) { Say("There are no tracked changes to turn down."); return; }
        _active.Dirty = true;
        _active.Undo.Push(() => Restore(_active, before));
        Rebuild(_active);
        Say($"Turned down {settled} tracked change{(settled == 1 ? "" : "s")}: what was struck out is back.");
    }

    private void OnMore(object sender, RoutedEventArgs e)
    {
        if (sender is not Button button || button.ContextMenu is null) return;
        button.ContextMenu.PlacementTarget = button;
        button.ContextMenu.Placement = System.Windows.Controls.Primitives.PlacementMode.Bottom;
        button.ContextMenu.IsOpen = true;
    }

    /// <summary>
    /// Which version this is, and where the help lives. An IT department cannot support three hundred people if the
    /// first question on every call needs a command prompt to answer.
    /// </summary>
    private void OnAbout(object sender, RoutedEventArgs e)
    {
        MessageBox.Show(this,
            $"Plain for Windows {Cli.Version}\n\n" +
            "Opens Word, Excel and PowerPoint files, edits the basics, and never damages what it doesn't understand.\n\n" +
            "The panel on the right names everything in a file that Plain keeps but cannot draw. All of it is written " +
            "back exactly as it was found, so nothing you cannot see is at risk when you save.\n\n" +
            "Free and MIT licensed. No account, no cloud, nothing sent anywhere.\n\n" +
            $"Settings and kept copies live in:\n{Settings.Folder}",
            "About Plain", MessageBoxButton.OK, MessageBoxImage.Information);
    }

    private void OnSettings(object sender, RoutedEventArgs e)
    {
        var dialog = new ReadingSettings(_settings) { Owner = this };
        if (dialog.ShowDialog() != true) return;
        _settings.Save();
        ApplyTextScale();
        foreach (var file in _open) if (file.View is DocView doc) doc.ApplyReading(_settings);
        Refresh();
        Say("Reading settings saved.");
    }

    /// <summary>Copy the top cell of the selection down through the rest of it, the way a column of rates gets filled.</summary>
    private void FillDown()
    {
        if (_active?.View is not WorkbookView view) return;
        int filled = view.CurrentGrid.FillDown();
        Say(filled == 0 ? "Select a cell and the ones below it, then Ctrl+D fills them from the top one."
                        : $"Filled {filled} cell{(filled == 1 ? "" : "s")} from the one above.");
    }

    // ---------- find ----------

    private int _findAt = -1;
    private int _findCount;

    private void ShowFind(bool replacing = false)
    {
        if (_active is null) return;
        FindBar.Visibility = Visibility.Visible;
        // Either way the caret starts in what you are looking for; there is nothing to replace until that is typed.
        FindBox.Focus();
        FindBox.SelectAll();
        RunFind(FindBox.Text);
    }

    private void OnFindOptionChanged(object sender, RoutedEventArgs e) => RunFind(FindBox.Text);

    /// <summary>
    /// Change every occurrence at once. This is the single thing the panels asked for most, and it is one step of
    /// undo however many it changed, because a replace-all you cannot take back is a frightening thing to click.
    /// </summary>
    private void OnReplaceAll(object sender, RoutedEventArgs e)
    {
        if (_active is null || FindBox.Text.Length == 0) return;

        var options = new Replace.Options(MatchCase: MatchCase.IsChecked == true);
        var before = Snapshot(_active);
        var result = Replace.InFile(_active.File, FindBox.Text, ReplaceBox.Text, options);

        if (!result.Any) { FindCount.Text = "not found"; Say($"\"{FindBox.Text}\" is not in this file."); return; }

        _active.Dirty = true;
        _active.Undo.Push(() => Restore(_active, before));
        Rebuild(_active);
        Say($"Changed {result.Occurrences} occurrence{(result.Occurrences == 1 ? "" : "s")} in {result.Cells} place{(result.Cells == 1 ? "" : "s")}.");
        RunFind(FindBox.Text);
    }

    /// <summary>
    /// Everything a whole-file change touches, kept so it can be put back. Replace-all can reach hundreds of cells
    /// across several sheets, so the only honest undo is the state that was there before.
    /// </summary>
    private static Dictionary<string, byte[]> Snapshot(OpenFile file)
    {
        file.File.Flush();
        var kept = new Dictionary<string, byte[]>(StringComparer.Ordinal);
        foreach (var part in file.File.ShownParts())
            if (file.File.Package.Has(part)) kept[part] = file.File.Package.Read(part);
        return kept;
    }

    private static void Restore(OpenFile file, Dictionary<string, byte[]> kept)
    {
        foreach (var (part, bytes) in kept) file.File.Package.Write(part, bytes);
    }

    /// <summary>Build the view again from the file, after a change too broad to patch in place.</summary>
    private void Rebuild(OpenFile file)
    {
        // Push what the models are holding into the package first. Replace-all and inserting a row change the
        // workbook, not the package, so reading the package without this gave back the file as it was before, and
        // the change was announced and then quietly thrown away.
        file.File.Flush();

        // The rebuilt file must know where it came from, or the next save has nowhere to go.
        var reopened = PlainFile.Read(file.File.Package.ToBytes(), file.FilePath);
        var replacement = Build(reopened, file.FilePath);
        replacement.Dirty = true;
        foreach (var step in file.Undo.Reverse()) replacement.Undo.Push(step);
        int at = _open.IndexOf(file);
        if (at >= 0) _open[at] = replacement;
        if (_active == file) _active = replacement;
        Refresh();
    }

    private void OnFindClose(object sender, RoutedEventArgs e)
    {
        FindBar.Visibility = Visibility.Collapsed;
        (_active?.View as UIElement)?.Focus();
    }

    private void OnFindChanged(object sender, TextChangedEventArgs e) => RunFind(FindBox.Text);

    private void RunFind(string term)
    {
        if (_active?.View is not IFindable findable) { FindCount.Text = ""; return; }
        _findCount = findable.FindAll(term);
        _findAt = -1;
        FindCount.Text = term.Length == 0 ? "" : _findCount == 0 ? "not found" : $"{_findCount} found";
        FindPrev.IsEnabled = FindNext.IsEnabled = _findCount > 0;
        if (_findCount > 0) Step(1);
    }

    private void Step(int by)
    {
        if (_active?.View is not IFindable findable || _findCount == 0) return;
        _findAt = ((_findAt + by) % _findCount + _findCount) % _findCount;   // wrap both ways
        string where = findable.Reveal(_findAt);
        FindCount.Text = $"{_findAt + 1} of {_findCount}" + (where.Length > 0 ? $"  ·  {where}" : "");
    }

    private void OnFindNext(object sender, RoutedEventArgs e) => Step(1);
    private void OnFindPrevious(object sender, RoutedEventArgs e) => Step(-1);

    private void OnFindKey(object sender, KeyEventArgs e)
    {
        if (e.Key == Key.Escape) { OnFindClose(sender, e); e.Handled = true; }
        else if (e.Key == Key.Enter)
        {
            Step(Keyboard.Modifiers == ModifierKeys.Shift ? -1 : 1);
            e.Handled = true;
        }
    }

    // ---------- the formula bar ----------

    private void OnFormulaKey(object sender, KeyEventArgs e)
    {
        if (_active?.View is not WorkbookView view) return;
        if (e.Key == Key.Enter) { view.Apply(CellEditor.Text); view.CurrentGrid.Focus(); e.Handled = true; }
        else if (e.Key == Key.Escape)
        {
            var grid = view.CurrentGrid;
            var cell = grid.Sheet.Read(grid.Selected);
            CellEditor.Text = cell.Formula ?? cell.Raw;
            grid.Focus();
            e.Handled = true;
        }
    }

    private void OnFormulaCommit(object sender, RoutedEventArgs e)
    {
        if (_active?.View is WorkbookView view && CellEditor.IsKeyboardFocusWithin == false) view.Apply(CellEditor.Text);
    }

    /// <summary>
    /// Put a row or column in, or take one out. Every formula in the workbook is rewritten to still mean what it
    /// meant, so this is too broad to patch in place: the view is built again from the file afterwards.
    /// </summary>
    private void ChangeGrid(OpenFile file, WorkbookView view, GridEdit edit, int at)
    {
        try
        {
            var before = Snapshot(file);
            int adjusted = file.File.Workbook!.Apply(view.CurrentSheet, edit, at);
            file.Dirty = true;
            file.Undo.Push(() => Restore(file, before));
            Rebuild(file);
            string what = edit switch
            {
                GridEdit.InsertRow => $"Put a row in at {at}",
                GridEdit.DeleteRow => $"Took row {at} out",
                GridEdit.InsertColumn => $"Put a column in at {CellRef.ColumnName(at)}",
                _ => $"Took column {CellRef.ColumnName(at)} out",
            };
            Say($"{what}. {adjusted} formula{(adjusted == 1 ? "" : "s")} adjusted to match.");
        }
        catch (Exception ex) { Say("Could not change the sheet: " + Explain(ex)); }
    }

    /// <summary>Write the file out as a PDF, which more panels asked for than anything else.</summary>
    private void OnPdf(object sender, RoutedEventArgs e)
    {
        if (_active is null) return;
        var dialog = new SaveFileDialog
        {
            Title = "Save as PDF",
            FileName = Path.GetFileNameWithoutExtension(_active.FilePath) + ".pdf",
            DefaultExt = ".pdf",
            Filter = "PDF|*.pdf",
            InitialDirectory = Path.GetDirectoryName(_active.FilePath),
        };
        if (dialog.ShowDialog(this) != true) return;
        try
        {
            _active.File.Flush();
            var result = PdfExport.Build(_active.File, Path.GetFileNameWithoutExtension(_active.FilePath));
            File.WriteAllBytes(dialog.FileName, result.Bytes);
            Say($"Wrote {Path.GetFileName(dialog.FileName)}: {result.Pages} page{(result.Pages == 1 ? "" : "s")}." +
                (result.Warning is null ? "" : "  " + result.Warning));
        }
        catch (Exception ex) { Say("Could not write the PDF: " + Explain(ex)); }
    }

    private void OnBullet(object sender, RoutedEventArgs e) => List(bulleted: true);
    private void OnNumberList(object sender, RoutedEventArgs e) => List(bulleted: false);

    private void List(bool bulleted)
    {
        if (_active?.View is not DocView doc || _active.File.Document is null) return;
        if (doc.FocusedBlock is not { } index) { Say("Click the line you want to change first."); return; }
        try
        {
            var before = Snapshot(_active);
            bool on = !_active.File.Document.IsList(index);
            if (!_active.File.Document.SetList(index, bulleted, on))
            {
                Say("This document carries no list numbering, and Plain will not invent one that would not match it.");
                return;
            }
            _active.Dirty = true;
            _active.Undo.Push(() => Restore(_active, before));
            Rebuild(_active);
            Say(on ? $"That line is a {(bulleted ? "bullet" : "numbered item")} now." : "That line is out of the list.");
        }
        catch (Exception ex) { Say("Could not change that: " + Explain(ex)); }
    }

    // ---------- bold, italic, headings ----------

    private void OnBold(object sender, RoutedEventArgs e) => Mark("b");
    private void OnItalic(object sender, RoutedEventArgs e) => Mark("i");

    /// <summary>Turn bold or italic on for whatever is selected, or off if it is already on throughout.</summary>
    private void Mark(string mark)
    {
        if (_active is null) return;
        try
        {
            var before = Snapshot(_active);
            switch (_active.View)
            {
                case WorkbookView view:
                {
                    var cells = view.CurrentGrid.SelectedCells().ToList();
                    if (cells.Count == 0) return;
                    bool on = !cells.All(c => view.CurrentGrid.Sheet.HasWeight(c, mark));
                    view.CurrentGrid.Sheet.SetWeight(cells, mark == "b" ? on : null, mark == "i" ? on : null);
                    view.CurrentGrid.Reload();
                    Say($"{cells.Count} cell{(cells.Count == 1 ? "" : "s")} {(on ? "now" : "no longer")} {(mark == "b" ? "bold" : "italic")}.");
                    break;
                }
                case DocView doc:
                {
                    int? at = doc.FocusedBlock;
                    if (at is not { } index) { Say("Click the line you want to change first."); return; }
                    bool on = !_active.File.Document!.IsAll(index, mark);
                    _active.File.Document.SetMark(index, mark, on);
                    Say($"That line is {(on ? "now" : "no longer")} {(mark == "b" ? "bold" : "italic")}.");
                    break;
                }
                default: return;
            }
            _active.Dirty = true;
            _active.Undo.Push(() => Restore(_active, before));
        }
        catch (Exception ex) { Say("Could not change that: " + Explain(ex)); }
        Refresh();
    }

    private bool _settingStyle;

    private void OnStylePicked(object sender, SelectionChangedEventArgs e)
    {
        if (_settingStyle || _active?.View is not DocView doc) return;
        if (StylePicker.SelectedItem is not ComboBoxItem { Tag: BlockKind kind }) return;
        if (doc.FocusedBlock is not { } index) { Say("Click the line you want to change first."); return; }

        try
        {
            var before = Snapshot(_active);
            if (!_active.File.Document!.SetKind(index, kind))
            {
                Say("This document has no heading style of that level, and Plain will not invent one.");
                return;
            }
            _active.Dirty = true;
            _active.Undo.Push(() => Restore(_active, before));
            Rebuild(_active);
            Say(kind == BlockKind.Paragraph ? "That line is ordinary text now." : $"That line is a {Name(kind)} now.");
        }
        catch (Exception ex) { Say("Could not change that: " + Explain(ex)); }
    }

    private static string Name(BlockKind kind) => kind switch
    {
        BlockKind.Heading1 => "main heading",
        BlockKind.Heading2 => "heading",
        BlockKind.Heading3 => "small heading",
        _ => "ordinary text",
    };

    // ---------- how cells show their numbers ----------

    private bool _settingFormat;

    private void OnFormatPicked(object sender, SelectionChangedEventArgs e)
    {
        if (_settingFormat || _active?.View is not WorkbookView view) return;
        if (FormatPicker.SelectedItem is not ComboBoxItem { Tag: string code }) return;

        var cells = view.CurrentGrid.SelectedCells().ToList();
        if (cells.Count == 0) return;
        try
        {
            var before = Snapshot(_active);
            view.CurrentGrid.Sheet.SetFormat(cells, code);
            _active.Dirty = true;
            _active.Undo.Push(() => Restore(_active, before));
            view.CurrentGrid.Reload();
            Say($"{cells.Count} cell{(cells.Count == 1 ? "" : "s")} now shown as {((ComboBoxItem)FormatPicker.SelectedItem).Content}.");
        }
        catch (Exception ex) { Say("Could not change the format: " + Explain(ex)); }
        Refresh();
    }

    // ---------- what other people wrote ----------

    /// <summary>One line of the comments panel, with what can be done to it.</summary>
    private sealed record NoteLine(string Who, string When, string What, string Id, string KeepLabel, string DropLabel)
    {
        public Visibility Actions => Id.Length > 0 ? Visibility.Visible : Visibility.Collapsed;
    }

    private void OnAcceptChanges(object sender, RoutedEventArgs e)
    {
        if (_active is null) return;
        var before = Snapshot(_active);
        int settled = Annotations.AcceptRevisions(_active.File);
        if (settled == 0) { Say("There are no tracked changes to settle."); return; }
        _active.Dirty = true;
        _active.Undo.Push(() => Restore(_active, before));
        Rebuild(_active);
        Say($"Settled {settled} tracked change{(settled == 1 ? "" : "s")}: what was added stayed, what was struck out went.");
    }

    private void OnRemoveComments(object sender, RoutedEventArgs e)
    {
        if (_active is null) return;
        var before = Snapshot(_active);
        int removed = Annotations.RemoveComments(_active.File);
        if (removed == 0) { Say("There are no comments to remove."); return; }
        _active.Dirty = true;
        _active.Undo.Push(() => Restore(_active, before));
        Rebuild(_active);
        Say($"Removed {removed} comment{(removed == 1 ? "" : "s")}.");
    }

    // ---------- printing ----------

    /// <summary>
    /// Print what is on screen. Plain does not lay pages out the way Word does, so this prints the view rather than
    /// claiming to reproduce someone else's pagination, and the dialog says as much.
    /// </summary>
    private void OnPrint(object sender, RoutedEventArgs e)
    {
        if (_active?.View is not FrameworkElement view) return;

        var dialog = new System.Windows.Controls.PrintDialog();
        if (dialog.ShowDialog() != true) return;

        try
        {
            double width = dialog.PrintableAreaWidth, height = dialog.PrintableAreaHeight;
            var document = Printing.Build(_active.File, _active.Name, width, height);
            dialog.PrintDocument(((System.Windows.Documents.IDocumentPaginatorSource)document).DocumentPaginator,
                                 _active.Name);
            Say($"Sent {_active.Name} to {dialog.PrintQueue?.Name ?? "the printer"}.");
        }
        catch (Exception ex) { Say("Could not print: " + Explain(ex)); }
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
        CopyBtn.IsEnabled = _active is not null;
        UndoBtn.IsEnabled = _active?.Undo.Count > 0;
        SaveBtn.Content = _active?.Dirty == true ? "Save" : "Saved";

        ContextHint.Text = _active?.File.Kind switch
        {
            FileKind.Spreadsheet => "Click a cell and type. A leading = makes a formula. Shift to select a block, then copy or paste.",
            FileKind.Document => "Click any line and type. One column, no page breaks.",
            FileKind.Presentation => "Pick a slide, then edit its text.",
            _ => "",
        };

        MoreBtn.IsEnabled = true;
        FormatPicker.Visibility = _active?.File.Kind == FileKind.Spreadsheet ? Visibility.Visible : Visibility.Collapsed;

        if (_active is null)
        {
            RailList.ItemsSource = null;
            RailFoot.Text = "";
            KeepBtn.Content = "Preserved";
            NotesBtn.Visibility = Visibility.Collapsed;
            Notes.Visibility = Visibility.Collapsed;
            StatusPromise.Text = _message.Length > 0 ? _message : "Nothing open.";
            StatusStat.Text = "";
            Title = "Plain";

            var recent = _settings.RecentThatExist()
                .Select(p => new RecentEntry(Path.GetFileName(p), p)).ToList();
            RecentList.ItemsSource = recent;
            RecentHeading.Visibility = recent.Count > 0 ? Visibility.Visible : Visibility.Collapsed;
            return;
        }

        FillFormatPicker();
        FillStylePicker();
        FillNotes();
        FillPictures();
        bool canMark = _active.File.Kind is FileKind.Spreadsheet or FileKind.Document;
        BoldBtn.Visibility = ItalicBtn.Visibility = canMark ? Visibility.Visible : Visibility.Collapsed;

        bool document = _active.File.Kind == FileKind.Document;
        BulletBtn.Visibility = NumberBtn.Visibility = document ? Visibility.Visible : Visibility.Collapsed;
        if (document && _active.File.Document is not null)
        {
            var (bullets, numbers) = _active.File.Document.ListsAvailable();
            BulletBtn.IsEnabled = bullets;
            NumberBtn.IsEnabled = numbers;
        }

        var parts = _active.File.Parts();
        var kept = parts.Where(p => p.Role == PartRole.Preserved).ToList();
        var rows = Preserved.Summarise(parts);
        var (bookkeeping, bookkeepingBytes) = Preserved.Bookkeeping(parts);

        RailList.ItemsSource = rows;
        KeepBtn.Content = $"Preserved  {kept.Count}";

        // One line for the plumbing, rather than a row each for things nobody has ever wanted to look at.
        string plumbing = bookkeeping == 0 ? ""
            : $"Plus {bookkeeping} parts of bookkeeping the file needs to be a file: what points at what, which part is which, the settings and theme it was made with. Kept, not listed.\n\n";

        RailFoot.Text = plumbing + _active.File.Kind switch
        {
            FileKind.Spreadsheet => "Change a cell and the totals that read it lose their stored value until Excel works them out again. Every other total keeps its number.",
            FileKind.Document => "Tracked changes, comments, headers and footers stay in the file. Plain shows the text of the body.",
            _ => "Shapes Plain cannot draw are held in place. Editing a title never moves them.",
        };

        Notes.Visibility = _notesVisible ? Visibility.Visible : Visibility.Collapsed;
        Pictures.Visibility = _picturesVisible ? Visibility.Visible : Visibility.Collapsed;

        if (rows.Count == 0 && bookkeeping > 0)
            RailFoot.Text = $"Everything in this file is something Plain shows, apart from {bookkeeping} parts of bookkeeping. Nothing is being held back.";

        int shown = parts.Count(p => p.Role == PartRole.Shown);
        StatusPromise.Text = _message.Length > 0
            ? _message
            : $"{parts.Count} parts read, {shown} shown, {parts.Count - shown} kept byte for byte";

        StatusStat.Text = _active.File.Kind switch
        {
            FileKind.Spreadsheet => Describe(_active),
            FileKind.Document => $"{Counts.Of(_active.File.Document!.PlainText()).Words:N0} words, flowing view with no page breaks",
            _ => $"Slide {(_active.View as DeckView)?.Current.Number} of {_active.File.Deck!.Slides.Count}",
        };

        if (_active.Flattened)
            StatusStat.Text += "  ·  one block's mixed formatting was flattened";

        Title = (_active.Dirty ? "• " : "") + _active.Name + " - Plain";
    }

    /// <summary>Offer the formats by name, showing the one the selected cell already uses.</summary>
    private void FillFormatPicker()
    {
        if (_active?.View is not WorkbookView view) return;
        _settingFormat = true;
        try
        {
            if (FormatPicker.Items.Count == 0)
                foreach (var (name, code) in Styles.Common)
                    FormatPicker.Items.Add(new ComboBoxItem { Content = name, Tag = code });

            string current = view.CurrentGrid.Sheet.FormatOf(view.CurrentGrid.Selected);
            var match = FormatPicker.Items.Cast<ComboBoxItem>().FirstOrDefault(i => (string)i.Tag == current);
            FormatPicker.SelectedItem = match;
            if (match is null) FormatPicker.Text = current.Length == 0 ? "General" : current;
        }
        finally { _settingFormat = false; }
    }

    /// <summary>Offer the heading levels this document actually has.</summary>
    private void FillStylePicker()
    {
        if (_active?.File.Kind != FileKind.Document || _active.File.Document is null)
        {
            StylePicker.Visibility = Visibility.Collapsed;
            return;
        }
        StylePicker.Visibility = Visibility.Visible;
        _settingStyle = true;
        try
        {
            var kinds = _active.File.Document.AvailableKinds();
            if (StylePicker.Items.Count != kinds.Count)
            {
                StylePicker.Items.Clear();
                foreach (var kind in kinds)
                    StylePicker.Items.Add(new ComboBoxItem { Content = Name(kind), Tag = kind });
            }
            var at = (_active.View as DocView)?.FocusedBlock;
            var current = at is { } index ? _active.File.Document.Read(index).Kind : BlockKind.Paragraph;
            if (current is BlockKind.ListItem or BlockKind.TableCell) current = BlockKind.Paragraph;
            StylePicker.SelectedItem = StylePicker.Items.Cast<ComboBoxItem>()
                .FirstOrDefault(i => (BlockKind)i.Tag == current);
        }
        finally { _settingStyle = false; }
    }

    /// <summary>One line of the pictures panel, with the picture itself where it can be drawn.</summary>
    private sealed record PictureLine(string Kind, string Size, string Part, System.Windows.Media.ImageSource? Preview);

    /// <summary>The pictures in the file, drawn where Windows can draw them and listed where it cannot.</summary>
    private void FillPictures()
    {
        if (_active is null) return;
        var pictures = Media.In(_active.File);
        PicturesBtn.Visibility = pictures.Count > 0 ? Visibility.Visible : Visibility.Collapsed;
        PicturesBtn.Content = $"Pictures  {pictures.Count}";
        if (pictures.Count == 0) { _picturesVisible = false; return; }

        var lines = new List<PictureLine>();
        foreach (var picture in pictures)
        {
            System.Windows.Media.ImageSource? preview = null;
            if (Media.Drawable(picture.Part))
            {
                try
                {
                    var image = new System.Windows.Media.Imaging.BitmapImage();
                    image.BeginInit();
                    image.CacheOption = System.Windows.Media.Imaging.BitmapCacheOption.OnLoad;
                    image.StreamSource = new MemoryStream(picture.Read());
                    image.DecodePixelWidth = 260;      // no need for the full size in a narrow panel
                    image.EndInit();
                    image.Freeze();
                    preview = image;
                }
                catch { }   // a picture Windows will not decode is listed rather than shown
            }
            lines.Add(new PictureLine(picture.Kind + (preview is null ? " (not shown)" : ""), picture.Size, picture.Part, preview));
        }
        PictureList.ItemsSource = lines;
        PicturesLede.Text = $"{pictures.Count} picture{(pictures.Count == 1 ? "" : "s")} are in this file. Plain keeps them exactly as they are; this is so you can see what you are sending.";
    }

    /// <summary>What other people wrote, and whether there is anything to show at all.</summary>
    private void FillNotes()
    {
        if (_active is null) return;
        var comments = Annotations.Comments(_active.File);
        var revisions = Annotations.Revisions(_active.File);
        int total = comments.Count + revisions.Count;

        NotesBtn.Visibility = total > 0 ? Visibility.Visible : Visibility.Collapsed;
        NotesBtn.Content = $"Comments  {total}";
        if (total == 0) { _notesVisible = false; return; }

        var lines = new List<NoteLine>();
        bool canAct = _active.File.Kind == FileKind.Document;
        foreach (var note in comments)
            lines.Add(new NoteLine(note.Author.Length > 0 ? note.Author : "Someone",
                                   note.When,
                                   note.Text + (note.Where.Length > 0 && note.Where != "comment" ? $"  ({note.Where})" : ""),
                                   canAct ? $"comment:{note.Index}" : "",
                                   "Keep", "Remove"));
        foreach (var revision in revisions)
            lines.Add(new NoteLine(revision.Author.Length > 0 ? revision.Author : "Someone",
                                   revision.When,
                                   (revision.Inserted ? "Added: " : "Struck out: ") + revision.Text,
                                   canAct ? $"change:{revision.Index}" : "",
                                   "Accept", "Turn down"));
        NotesList.ItemsSource = lines;

        NotesLede.Text = $"{comments.Count} comment{(comments.Count == 1 ? "" : "s")} and " +
                         $"{revisions.Count} tracked change{(revisions.Count == 1 ? "" : "s")} are in this file. " +
                         "They stay in it whether or not you look at them.";
        AcceptBtn.IsEnabled = RejectBtn.IsEnabled = revisions.Count > 0;
        ClearNotesBtn.IsEnabled = comments.Count > 0;
    }

    private static string Describe(OpenFile file)
    {
        if (file.View is not WorkbookView view) return "";
        var sheet = view.CurrentSheet;
        int count = file.File.Workbook!.Sheets.Count;
        string where = count > 1 ? $"{sheet.Name} of {count} sheets" : sheet.Name;
        return $"{where}, used to {sheet.Extent}";
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
