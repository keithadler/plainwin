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

/// <summary>Pushing a step that only knows how to undo itself, which most of them are.</summary>
internal static class UndoStack
{
    public static void Push(this Stack<Edit> stack, Action undo) => stack.Push(new Edit(undo, null));
}

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
        public Stack<Edit> Undo { get; } = new();

        /// <summary>Steps that have been undone and can be done again, emptied the moment something new is typed.</summary>
        public Stack<Edit> Redo { get; } = new();
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
                StartUpdateCheck();
            }
            Refresh();
        };
        KeyDown += OnWindowKey;
    }

    /// <summary>Put the grid on a cell, so a picture can show a sheet part way down rather than always at A1.</summary>
    internal void ScrollForScreenshot(string cell)
    {
        if (_active?.View is WorkbookView book && Core.CellRef.TryParse(cell, out var reference))
            book.CurrentGrid.Select(reference);
    }

    /// <summary>Open a file straight away, for the screenshot renderer which has no user to click Open.</summary>
    internal void OpenForScreenshot(string path, bool showNotes = false)
    {
        OpenPath(path);
        _notesVisible = showNotes;
        _railVisible = !showNotes;   // the pictures show the app as it is meant to look, not as a setting left it
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
            CommitPendingEdit();
            foreach (var file in _open.Where(f => f.Dirty))
                Recovery.Keep(file.File, file.FilePath);
        };
        _keeper.Start();
    }

    /// <summary>
    /// Ask GitHub once a day whether there is a newer version. Everything about this is deliberately quiet: it runs
    /// in the background, it never blocks opening a file, a failure says nothing at all, and the most it does when it
    /// finds something is put a sentence in the status bar with a button beside it. Nothing downloads or installs
    /// itself, and the switch to stop it is in Reading and settings.
    /// </summary>
    private void StartUpdateCheck()
    {
        UpdateCheck.Start(_settings, found =>
        {
            _updateFound = found;
            UpdateBar.Visibility = Visibility.Visible;
            UpdateText.Text = $"Version {found.Version} is out. You have {Cli.Version}.";
        });
    }

    private Core.Updates.Available? _updateFound;

    private void OnUpdateOpen(object sender, RoutedEventArgs e)
    {
        if (_updateFound is not null) UpdateCheck.OpenReleasePage(_updateFound.Page);
    }

    private void OnUpdateDismiss(object sender, RoutedEventArgs e) => UpdateBar.Visibility = Visibility.Collapsed;

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
    /// <summary>Ask which of the three to make. The kind is a choice, not something buried in a file dialog.</summary>
    private void OnNew(object sender, RoutedEventArgs e)
    {
        if (sender is not Button button || button.ContextMenu is null) { MakeNew(FileKind.Spreadsheet); return; }
        button.ContextMenu.PlacementTarget = button;
        button.ContextMenu.Placement = System.Windows.Controls.Primitives.PlacementMode.Bottom;
        button.ContextMenu.IsOpen = true;
    }

    private void OnNewSheet(object sender, RoutedEventArgs e) => MakeNew(FileKind.Spreadsheet);
    private void OnNewDoc(object sender, RoutedEventArgs e) => MakeNew(FileKind.Document);
    private void OnNewDeck(object sender, RoutedEventArgs e) => MakeNew(FileKind.Presentation);

    /// <summary>
    /// Make an empty file of the chosen kind and open it. It is written to disk before you type into it, so there is
    /// nothing to lose if the machine gives up half way through your first paragraph.
    /// </summary>
    private void MakeNew(FileKind kind)
    {
        var (extension, what, filter) = kind switch
        {
            FileKind.Document => (".docx", "Document", "Word document|*.docx"),
            FileKind.Presentation => (".pptx", "Presentation", "PowerPoint deck|*.pptx"),
            _ => (".xlsx", "Spreadsheet", "Excel workbook|*.xlsx"),
        };

        var dialog = new SaveFileDialog
        {
            Title = $"New {what.ToLowerInvariant()}",
            FileName = "Untitled" + extension,
            DefaultExt = extension,
            Filter = filter + "|All files|*.*",
            OverwritePrompt = false,   // Plain refuses to write over one itself, with a clearer message
        };
        if (dialog.ShowDialog(this) != true) return;

        // Whatever they name it, the extension decides the kind, so a typed .docx still gets a document.
        var path = dialog.FileName;
        if (PlainFile.KindOf(path) == FileKind.Unknown) path += extension;

        try
        {
            var made = PlainFile.Create(path);
            var entry = Build(made, path);
            _open.Add(entry);
            _active = entry;
            _settings.Remember(path);
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
                bookView.Edited += edit => { entry.Undo.Push(edit); entry.Redo.Clear(); entry.Dirty = true; Refresh(); };
                bookView.GridChangeRequested += (edit, at) => ChangeGrid(entry, bookView, edit, at);
                bookView.SortRequested += (t, b, l, r, key, up) => SortRows(entry, bookView, t, b, l, r, key, up);
                bookView.FilterRequested += (column, row) => FilterRows(bookView, column, row);
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
        // Control, and Control with Shift, which is how redo is spelled. Anything with Alt or Windows is not ours.
        var held = e.KeyboardDevice.Modifiers;
        if ((held & ModifierKeys.Control) == 0) return;
        if ((held & (ModifierKeys.Alt | ModifierKeys.Windows)) != 0) return;
        switch (e.Key)
        {
            case Key.O: OnOpen(sender, e); e.Handled = true; break;
            case Key.S: OnSave(sender, e); e.Handled = true; break;
            case Key.Z:
                // Ctrl+Shift+Z is redo on every other program, and so is Ctrl+Y.
                if ((Keyboard.Modifiers & ModifierKeys.Shift) != 0) OnRedo(sender, e); else OnUndo(sender, e);
                e.Handled = true; break;
            case Key.Y: OnRedo(sender, e); e.Handled = true; break;
            case Key.G: GoToCell(); e.Handled = true; break;
            case Key.W: CloseActive(); e.Handled = true; break;
            case Key.F:
                if ((Keyboard.Modifiers & ModifierKeys.Shift) != 0) OnFindInFolder(sender, e); else ShowFind();
                e.Handled = true; break;
            case Key.H: ShowFind(replacing: true); e.Handled = true; break;
            case Key.N: MakeNew(FileKind.Spreadsheet); e.Handled = true; break;
            case Key.P: OnPrint(sender, e); e.Handled = true; break;
            case Key.D: FillDown(); e.Handled = true; break;
            case Key.B: Mark("b"); e.Handled = true; break;
            case Key.I: Mark("i"); e.Handled = true; break;
            case Key.OemPlus or Key.Add: Bigger(); e.Handled = true; break;
            case Key.OemMinus or Key.Subtract: Smaller(); e.Handled = true; break;
            case Key.D0 or Key.NumPad0: NormalSize(); e.Handled = true; break;
        }
    }

    /// <summary>
    /// The editors hand their text over when they lose the caret, which is what keeps a long document from rebuilding
    /// on every keystroke. Anything that reads the file has to take the caret away first, or it reads the file without
    /// the words just typed. A file with one box in it, which is exactly what a new blank document is, never loses the
    /// caret on its own, so without this you could type a page into a new document, save, and save nothing.
    /// </summary>
    private void CommitPendingEdit()
    {
        if (Keyboard.FocusedElement is not TextBox box) return;
        box.GetBindingExpression(TextBox.TextProperty)?.UpdateSource();
        var scope = FocusManager.GetFocusScope(box);
        FocusManager.SetFocusedElement(scope, null);
        Keyboard.ClearFocus();
        // Put the caret back where the typing was, so saving does not also move you.
        if (box.IsVisible) { int at = box.SelectionStart; box.Focus(); box.Select(at, 0); }
    }

    /// <summary>
    /// Sort the selected rows. Plain refuses whenever a formula would be made to mean something else, and when it
    /// refuses it says why in the status bar rather than doing something almost right. The whole block is kept for
    /// undo before anything moves, so Ctrl+Z puts the rows back exactly as they were.
    /// </summary>
    private void SortRows(OpenFile entry, WorkbookView view, int top, int bottom, int left, int right,
                          int keyColumn, bool ascending)
    {
        var book = entry.File.Workbook;
        var sheet = view.CurrentSheet;
        if (book is null || sheet is null) return;

        // Remember what was there, so undo is exact rather than a second sort in the other direction.
        var before = new List<(CellRef At, string Value)>();
        for (int r = top; r <= bottom; r++)
            for (int c = left; c <= right; c++)
            {
                var cell = sheet.Read(new CellRef(c, r));
                before.Add((new CellRef(c, r),
                    cell.Kind is Core.CellKind.Number or Core.CellKind.Boolean ? cell.Raw
                    : cell.Kind == Core.CellKind.Empty ? "" : cell.Display));
            }

        var result = Core.Sort.Rows(book, sheet, top, bottom, left, right, keyColumn, ascending);
        if (result is Core.Sort.Refused refused) { Say(refused.Reason); return; }

        var sorted = (Core.Sort.Sorted)result;
        entry.Undo.Push(() =>
        {
            foreach (var (at, value) in before) sheet.Set(at, value);
            view.Redraw();
        });
        entry.Dirty = true;
        view.Redraw();
        Say(sorted.RowsMoved == 0
            ? "Those rows were already in that order."
            : $"Sorted {bottom - top + 1} rows by column {Core.CellRef.ColumnName(keyColumn)}. {sorted.RowsMoved} moved.");
        Refresh();
    }

    private void OnSave(object sender, RoutedEventArgs e)
    {
        if (_active is null) return;
        CommitPendingEdit();

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
        CommitPendingEdit();
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
        var step = _active.Undo.Pop();
        step.Undo();

        // A step that cannot say how to repeat itself makes everything after it unrepeatable too, so the pile of
        // things waiting to be redone is thrown away rather than left to put things back in the wrong order.
        if (step.CanRedo) _active.Redo.Push(step); else _active.Redo.Clear();

        _active.Dirty = true;    // whatever is on disk, the file in front of you has just changed again
        Refresh();
    }

    /// <summary>
    /// Ctrl+G: ask for a cell and go there. A sheet with ten thousand rows needs a way to get to one of them that
    /// is not scrolling, and the answer everyone already knows is to type its name.
    /// </summary>
    private void GoToCell()
    {
        if (_active?.View is not WorkbookView book) return;

        var asked = Prompt("Go to", "Which cell? For example B14.", CellRefText.Text);
        if (asked is null) return;
        if (!book.GoTo(asked)) Say($"\"{asked}\" is not a cell. A cell is a letter and a number, like B14.");
    }

    /// <summary>
    /// Show only the rows whose cell in this column contains what you type. Nothing is written to the file: this is
    /// a way of looking at the sheet, and clearing it puts every row back.
    /// </summary>
    private void FilterRows(WorkbookView view, int column, int firstRow)
    {
        if (view.Filtering) { view.ClearFilter(); Say("Showing every row again."); Refresh(); return; }

        var asked = Prompt("Show only some rows",
            $"Show only rows where column {Core.CellRef.ColumnName(column)} contains:", "");
        if (asked is null) return;

        view.Filter(column, asked, firstRow);
        Say(view.Filtering ? view.FilterSaid : "Nothing matched, so every row is still showing.");
        Refresh();
    }

    /// <summary>A one-line question, because a whole dialog file for one box is more than this needs.</summary>
    private string? Prompt(string title, string question, string initial)
    {
        var box = new TextBox { Text = initial, Margin = new Thickness(0, 10, 0, 0), Padding = new Thickness(6, 4, 6, 4) };
        var ok = new Button { Content = "Go", IsDefault = true, MinWidth = 76, Margin = new Thickness(0, 14, 8, 0) };
        var cancel = new Button { Content = "Cancel", IsCancel = true, MinWidth = 76, Margin = new Thickness(0, 14, 0, 0) };
        var buttons = new StackPanel { Orientation = Orientation.Horizontal, HorizontalAlignment = HorizontalAlignment.Right };
        buttons.Children.Add(ok); buttons.Children.Add(cancel);

        var stack = new StackPanel { Margin = new Thickness(18) };
        stack.Children.Add(new TextBlock { Text = question, TextWrapping = TextWrapping.Wrap });
        stack.Children.Add(box);
        stack.Children.Add(buttons);

        var window = new Window
        {
            Title = title, Content = stack, Owner = this, SizeToContent = SizeToContent.Height, Width = 360,
            WindowStartupLocation = WindowStartupLocation.CenterOwner, ResizeMode = ResizeMode.NoResize,
            Background = App.B("Chrome"), Foreground = App.B("Ink"),
        };
        ok.Click += (_, _) => { window.DialogResult = true; };
        window.Loaded += (_, _) => { box.Focus(); box.SelectAll(); };
        return window.ShowDialog() == true && box.Text.Trim().Length > 0 ? box.Text.Trim() : null;
    }

    /// <summary>
    /// Which files in a folder hold the words. This only ever reads: it says which files and where, and then you
    /// open the ones that matter and change them yourself, one at a time, watching what happens. A tool that
    /// offered to change forty files at once would be asking for a great deal of trust for very little work saved.
    /// </summary>
    private void OnFindInFolder(object sender, RoutedEventArgs e)
    {
        var term = Prompt("Find in a folder", "Which words are you looking for?", "");
        if (term is null) return;

        var picker = new Microsoft.Win32.OpenFolderDialog { Title = "Which folder?" };
        if (picker.ShowDialog(this) != true) return;

        Say($"Looking through {picker.FolderName}…");
        var report = Core.Folder.Search(picker.FolderName, term, deep: true);

        var stack = new StackPanel { Margin = new Thickness(18) };
        stack.Children.Add(new TextBlock
        {
            Text = report.Hits.Count == 0
                ? $"Nothing in those {report.Looked} files holds \"{term}\"."
                : $"{report.Hits.Count} of {report.Looked} files hold \"{term}\". Nothing has been changed.",
            TextWrapping = TextWrapping.Wrap, FontWeight = FontWeights.SemiBold, Margin = new Thickness(0, 0, 0, 12),
        });

        foreach (var hit in report.Hits)
        {
            var open = new Button
            {
                Content = $"{Path.GetFileName(hit.Path)}  ({hit.Count})",
                HorizontalAlignment = HorizontalAlignment.Left,
                Style = (Style)FindResource("Flat"),
                Tag = hit.Path,
            };
            open.Click += (b, _) =>
            {
                if (((Button)b).Tag is string p) OpenPath(p);
                Refresh();
            };
            stack.Children.Add(open);
            foreach (var place in hit.Places)
                stack.Children.Add(new TextBlock
                {
                    Text = "    " + place, TextWrapping = TextWrapping.Wrap, FontSize = 11.5,
                    Foreground = App.B("Ink2"), Margin = new Thickness(0, 0, 0, 2),
                });
        }

        if (report.Troubles.Count > 0)
            stack.Children.Add(new TextBlock
            {
                Text = $"{report.Troubles.Count} could not be read: "
                     + string.Join(", ", report.Troubles.Take(4).Select(t => Path.GetFileName(t.Path))),
                TextWrapping = TextWrapping.Wrap, FontSize = 11.5, Foreground = App.B("Ink3"),
                Margin = new Thickness(0, 12, 0, 0),
            });

        var window = new Window
        {
            Title = "Find in a folder", Content = new ScrollViewer { Content = stack }, Owner = this,
            Width = 620, Height = 480, WindowStartupLocation = WindowStartupLocation.CenterOwner,
            Background = App.B("Chrome"), Foreground = App.B("Ink"),
        };
        window.ShowDialog();
        Say(report.Hits.Count == 0 ? "Nothing found." : $"{report.Hits.Count} files hold it. Nothing was changed.");
    }

    /// <summary>
    /// What travels with this file that you probably did not mean to send: who wrote it, comments, tracked
    /// changes, hidden sheets and rows, speaker notes. It lists what is there and takes out only what is ticked.
    /// It never claims to have made a file safe, because it can only find what it knows to look for.
    /// </summary>
    private void OnBeforeYouSend(object sender, RoutedEventArgs e)
    {
        if (_active is null) return;
        CommitPendingEdit();

        var found = Core.Hidden.Find(_active.File);
        if (found.Count == 0)
        {
            MessageBox.Show(this,
                "Plain found nothing in this file that you would not expect to send: no names in the properties, "
                + "no comments, no tracked changes, nothing hidden.\n\n"
                + "It can only find what it knows to look for, so this is not a promise that the file holds nothing.",
                "Before you send it", MessageBoxButton.OK, MessageBoxImage.Information);
            return;
        }

        var boxes = new List<(CheckBox Box, Core.Hidden.Finding Finding)>();
        var stack = new StackPanel { Margin = new Thickness(18) };
        stack.Children.Add(new TextBlock
        {
            Text = "These travel with the file. Tick what you want taken out.",
            TextWrapping = TextWrapping.Wrap, Margin = new Thickness(0, 0, 0, 12),
        });

        foreach (var finding in found)
        {
            var box = new CheckBox
            {
                Content = $"{finding.Kind}  —  {finding.What}",
                IsEnabled = finding.CanRemove,
                Margin = new Thickness(0, 0, 0, 6),
            };
            if (!finding.CanRemove)
            {
                box.Content = $"{finding.Kind}  —  {finding.What}   (Plain leaves this alone)";
                box.ToolTip = "Hidden sheets and rows are somebody's working, and speaker notes live in parts of "
                            + "their own. Plain lists them so you know, and does not tear them out.";
            }
            boxes.Add((box, finding));
            stack.Children.Add(box);
        }

        stack.Children.Add(new TextBlock
        {
            Text = "Plain can only find what it knows to look for. This is not a promise that nothing else is in there.",
            TextWrapping = TextWrapping.Wrap, FontSize = 11.5, Margin = new Thickness(0, 8, 0, 0),
            Foreground = App.B("Ink3"),
        });

        var ok = new Button { Content = "Take them out", IsDefault = true, MinWidth = 110, Margin = new Thickness(0, 16, 8, 0) };
        var cancel = new Button { Content = "Leave it", IsCancel = true, MinWidth = 90, Margin = new Thickness(0, 16, 0, 0) };
        var buttons = new StackPanel { Orientation = Orientation.Horizontal, HorizontalAlignment = HorizontalAlignment.Right };
        buttons.Children.Add(ok); buttons.Children.Add(cancel);
        stack.Children.Add(buttons);

        var window = new Window
        {
            Title = "Before you send it", Content = new ScrollViewer { Content = stack }, Owner = this,
            SizeToContent = SizeToContent.Height, Width = 520, MaxHeight = 560,
            WindowStartupLocation = WindowStartupLocation.CenterOwner, ResizeMode = ResizeMode.NoResize,
            Background = App.B("Chrome"), Foreground = App.B("Ink"),
        };
        ok.Click += (_, _) => { window.DialogResult = true; };
        if (window.ShowDialog() != true) return;

        var wanted = boxes.Where(b => b.Box.IsChecked == true).Select(b => b.Finding.Kind).ToList();
        if (wanted.Count == 0) { Say("Nothing was ticked, so nothing was taken out."); return; }

        var did = Core.Hidden.Remove(_active.File, wanted);
        if (did.Count == 0) { Say("There was nothing to take out."); return; }

        _active.Dirty = true;
        Say(string.Join(" ", did) + " Save the file to write it back.");
        Refresh();
    }

    // ---------- the shape of a deck, and of a table ----------

    /// <summary>
    /// Adding, removing and moving slides, and rows in a table. These change the shape of a file rather than its
    /// words, so each one is done, saved into the model, and then the view is built again from the file: the
    /// numbering of slides and paragraphs is different afterwards, and carrying on with the old numbering is how a
    /// tool ends up editing the wrong thing.
    /// </summary>
    private void AfterShapeChange(string said)
    {
        if (_active is null) return;
        _active.Dirty = true;
        _active.Undo.Clear();     // the old steps point at slides and rows that have moved
        _active.Redo.Clear();
        Say(said + " Undo cannot reach past a change of shape, so the steps before it have been let go.");

        // Build the view again from the file as it now is.
        var path = _active.FilePath;
        _active.File.Flush();
        var rebuilt = PlainFile.Read(_active.File.Package.ToBytes(), path);
        var replacement = Build(rebuilt, path);
        replacement.Dirty = true;
        int at = _open.IndexOf(_active);
        _open[at] = replacement;
        _active = replacement;
        Refresh();
    }

    private void OnAddSlide(object sender, RoutedEventArgs e) => Slide(deck => Core.Slides.Add(
        _active!.File.Package, deck.Current.Number));

    private void OnRemoveSlide(object sender, RoutedEventArgs e) => Slide(deck => Core.Slides.Remove(
        _active!.File.Package, deck.Current.Number));

    private void OnMoveSlideUp(object sender, RoutedEventArgs e) => Slide(deck => Core.Slides.Move(
        _active!.File.Package, deck.Current.Number, deck.Current.Number - 1));

    private void OnMoveSlideDown(object sender, RoutedEventArgs e) => Slide(deck => Core.Slides.Move(
        _active!.File.Package, deck.Current.Number, deck.Current.Number + 1));

    private void Slide(Func<DeckView, Core.Slides.Result> what)
    {
        if (_active?.View is not DeckView deck) return;
        var outcome = what(deck);
        if (outcome is Core.Slides.Refused refused) { Say(refused.Reason); return; }
        AfterShapeChange(((Core.Slides.Done)outcome).What);
    }

    private void OnAddTableRow(object sender, RoutedEventArgs e) => TableRow(add: true);
    private void OnRemoveTableRow(object sender, RoutedEventArgs e) => TableRow(add: false);

    private void TableRow(bool add)
    {
        if (_active?.File.Document is not { } doc) return;
        var shape = doc.TableShape();
        if (shape.Count == 0) { Say("There are no tables in this document."); return; }

        int table = 0;
        if (shape.Count > 1)
        {
            var which = Prompt(add ? "Add a row" : "Take a row out",
                $"Which table? There are {shape.Count}, counting from the top.", "1");
            if (which is null) return;
            if (!int.TryParse(which, out table) || table < 1 || table > shape.Count)
            { Say($"There is no table {which}."); return; }
            table--;
        }

        var asked = Prompt(add ? "Add a row" : "Take a row out",
            add ? $"After which row? That table has {shape[table]}. Nought puts it at the top."
                : $"Which row? That table has {shape[table]}.",
            add ? shape[table].ToString() : "1");
        if (asked is null) return;
        if (!int.TryParse(asked, out var row)) { Say($"\"{asked}\" is not a row number."); return; }

        var outcome = add ? doc.InsertRow(table, row) : doc.DeleteRow(table, row);
        if (outcome is Core.TableRows.Refused refused) { Say(refused.Reason); return; }
        AfterShapeChange(((Core.TableRows.Done)outcome).What);
    }

    private void OnRedo(object sender, RoutedEventArgs e)
    {
        if (_active is null || _active.Redo.Count == 0) return;
        var step = _active.Redo.Pop();
        step.Redo!();
        _active.Undo.Push(step);
        _active.Dirty = true;
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
            $"Plain for Windows {Cli.Version}\n" +
            "Built by Keith Adler.\n\n" +
            "Opens Word, Excel and PowerPoint files, edits the basics, and never damages what it doesn't understand.\n\n" +
            "The panel on the right names everything in a file that Plain keeps but cannot draw. All of it is written " +
            "back exactly as it was found, so nothing you cannot see is at risk when you save.\n\n" +
            "Free and MIT licensed. No account, no cloud, no telemetry. The only thing it sends is a daily\n" +
            "question to GitHub about whether there is a newer version, which you can turn off in settings.\n\n" +
            "More small apps like this one: keithadler.github.io\n" +
            "Source and issues: github.com/keithadler/plainwin\n\n" +
            $"Settings and kept copies live in:\n{Settings.Folder}",
            "About Plain for Windows", MessageBoxButton.OK, MessageBoxImage.Information);
    }

    private void OnSettings(object sender, RoutedEventArgs e)
    {
        var dialog = new ReadingSettings(_settings) { Owner = this };
        if (dialog.ShowDialog() != true) return;
        _settings.Save();
        ApplyTextScale();
        foreach (var file in _open) if (file.View is DocView doc) doc.ApplyReading(_settings);
        Refresh();
        // The registry line is worth reporting, because it changed something outside Plain's own folder.
        Say(dialog.ExplorerMessage ?? "Settings saved.");
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
        CommitPendingEdit();
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
            var result = PdfExport.Build(_active.File, Path.GetFileNameWithoutExtension(_active.FilePath), _settings.Page());
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
            Say(kind == BlockKind.Paragraph ? "That line is ordinary text now." : $"That line is a {KindName(kind)} now.");
        }
        catch (Exception ex) { Say("Could not change that: " + Explain(ex)); }
    }

    private static string KindName(BlockKind kind) => kind switch
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
        CommitPendingEdit();
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
        RedoBtn.IsEnabled = _active?.Redo.Count > 0;

        bool isDeck = _active?.View is DeckView;
        bool hasTables = _active?.File.Document?.TableShape().Count > 0;
        foreach (var item in new[] { AddSlideItem, RemoveSlideItem, MoveSlideUpItem, MoveSlideDownItem })
            item.Visibility = isDeck ? Visibility.Visible : Visibility.Collapsed;
        foreach (var item in new[] { AddTableRowItem, RemoveTableRowItem })
            item.Visibility = hasTables ? Visibility.Visible : Visibility.Collapsed;
        ShapeSeparator.Visibility = isDeck || hasTables ? Visibility.Visible : Visibility.Collapsed;
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
                    StylePicker.Items.Add(new ComboBoxItem { Content = KindName(kind), Tag = kind });
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

        // What the selection adds up to comes first, because it is the thing being looked at right now.
        var summary = view.Summary();
        return summary.Length > 0 ? $"{summary}   ·   {where}" : $"{where}, used to {sheet.Extent}";
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
