using System.Globalization;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Input;
using System.Windows.Media;
using Plain.Core;
// Plain.Core has its own Grid, for rows and columns of a sheet; in this file Grid means the WPF panel.
using Grid = System.Windows.Controls.Grid;

namespace Plain;

/// <summary>
/// The grid. It draws only the cells that fit on screen and scrolls with real scroll bars, so a sheet with a hundred
/// thousand rows opens as fast as one with ten. Column widths come from the file, so the sheet looks like the one
/// whoever made it laid out.
/// </summary>
public sealed class SheetView : Grid
{
    private const double RowHeight = 24;
    private const double HeaderHeight = 25;
    private const double RowHeaderWidth = 46;
    private const double MinColumnWidth = 26;

    private readonly Sheet _sheet;
    private readonly Surface _surface;
    private readonly Strip _columns;
    private readonly Strip _rows;
    private readonly ScrollBar _vertical = new() { Orientation = Orientation.Vertical, SmallChange = 1, LargeChange = 10 };
    private readonly ScrollBar _horizontal = new() { Orientation = Orientation.Horizontal, SmallChange = 1, LargeChange = 5 };
    private readonly TextBox _editor;

    private readonly Dictionary<CellRef, Cell> _cache = new();
    private CellRef _selected = new(1, 1);
    private CellRef _anchor = new(1, 1);   // the other corner of the selection; equal to _selected for one cell
    private int _firstColumn = 1, _firstRow = 1;
    private int _lastColumn = 40, _lastRow = 200;

    public event Action<CellRef, Cell>? SelectionChanged;
    public event Action<Edit>? Edited;

    /// <summary>Say an edit happened. A step that knows how to repeat itself passes redo as well.</summary>
    private void Raise(Action undo, Action? redo = null) => Edited?.Invoke(new Edit(undo, redo));

    public Sheet Sheet => _sheet;
    public CellRef Selected => _selected;

    /// <summary>The rectangle currently selected, which is one cell unless it was extended with Shift.</summary>
    public (int Left, int Top, int Right, int Bottom) Range => (
        Math.Min(_anchor.Column, _selected.Column), Math.Min(_anchor.Row, _selected.Row),
        Math.Max(_anchor.Column, _selected.Column), Math.Max(_anchor.Row, _selected.Row));

    public bool HasRange => _anchor != _selected;

    /// <summary>
    /// What is in the selection, said the way a person would say it: how many numbers, what they add up to, and
    /// the average. This is the question a spreadsheet is usually opened to answer, and answering it in the status
    /// bar means not having to type a formula into an empty cell and then delete it again.
    /// </summary>
    public string Summary()
    {
        var (left, top, right, bottom) = Range;
        long cells = (long)(right - left + 1) * (bottom - top + 1);
        // A whole-column selection is millions of cells; reading them all to say "0 numbers" helps nobody.
        if (cells > 200_000) return "";

        int numbers = 0, filled = 0;
        double total = 0, low = double.MaxValue, high = double.MinValue;
        for (int r = top; r <= bottom; r++)
            for (int c = left; c <= right; c++)
            {
                var cell = Get(new CellRef(c, r));
                if (cell.Kind == CellKind.Empty) continue;
                filled++;
                if (!double.TryParse(cell.Raw, NumberStyles.Float, CultureInfo.InvariantCulture, out var value)) continue;
                numbers++;
                total += value;
                if (value < low) low = value;
                if (value > high) high = value;
            }

        if (numbers == 0) return filled == 0 ? "" : $"{filled} filled";
        if (numbers == 1) return $"1 number, {Show(total)}";
        return $"{numbers} numbers, sum {Show(total)}, average {Show(total / numbers)}, "
             + $"lowest {Show(low)}, highest {Show(high)}";
    }

    /// <summary>Enough decimals to be useful, not so many that a status bar turns into a wall of digits.</summary>
    private static string Show(double value)
    {
        double rounded = Math.Round(value, 4);
        return rounded == Math.Floor(rounded) && Math.Abs(rounded) < 1e15
            ? ((long)rounded).ToString("#,##0", CultureInfo.CurrentCulture)
            : rounded.ToString("#,##0.####", CultureInfo.CurrentCulture);
    }

    public SheetView(Sheet sheet)
    {
        _sheet = sheet;
        var extent = sheet.Extent;
        _lastColumn = Math.Max(extent.Column + 6, 12);
        _lastRow = Math.Max(extent.Row + 40, 60);

        ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(RowHeaderWidth) });
        ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
        ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });
        RowDefinitions.Add(new RowDefinition { Height = new GridLength(HeaderHeight) });
        RowDefinitions.Add(new RowDefinition { Height = new GridLength(1, GridUnitType.Star) });
        RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });

        var corner = new Border { Background = App.B("Surface2"), BorderBrush = App.B("Line"), BorderThickness = new Thickness(0, 0, 1, 1) };
        SetRow(corner, 0); SetColumn(corner, 0); Children.Add(corner);

        _columns = new Strip(this, horizontal: true);
        SetRow(_columns, 0); SetColumn(_columns, 1); Children.Add(_columns);

        _rows = new Strip(this, horizontal: false);
        SetRow(_rows, 1); SetColumn(_rows, 0); Children.Add(_rows);

        var host = new Grid { ClipToBounds = true };
        SetRow(host, 1); SetColumn(host, 1); Children.Add(host);

        _surface = new Surface(this);
        host.Children.Add(_surface);

        _editor = new TextBox
        {
            Visibility = Visibility.Collapsed,
            HorizontalAlignment = HorizontalAlignment.Left,
            VerticalAlignment = VerticalAlignment.Top,
            BorderThickness = new Thickness(2),
            BorderBrush = App.B("Accent"),
            Background = App.B("Surface"),
            Foreground = App.B("Ink"),
            Padding = new Thickness(4, 0, 0, 0),
            FontSize = 13,
            Height = RowHeight + 2,
        };
        _editor.KeyDown += EditorKey;
        _editor.LostFocus += (_, _) => CommitEditor();
        host.Children.Add(_editor);

        SetRow(_vertical, 1); SetColumn(_vertical, 2); Children.Add(_vertical);
        SetRow(_horizontal, 2); SetColumn(_horizontal, 1); Children.Add(_horizontal);
        _vertical.Scroll += (_, _) => { _firstRow = Math.Max(FrozenRows + 1, (int)_vertical.Value + 1); Redraw(); };
        _horizontal.Scroll += (_, _) => { _firstColumn = (int)_horizontal.Value + 1; Redraw(); };

        Background = App.B("Surface");
        Focusable = true;
        _surface.MouseLeftButtonDown += SurfaceClick;
        _surface.MouseWheel += (_, e) => { _vertical.Value = Math.Clamp(_vertical.Value - e.Delta / 40.0, _vertical.Minimum, _vertical.Maximum); _firstRow = Math.Max(FrozenRows + 1, (int)_vertical.Value + 1); Redraw(); };
        KeyDown += OnKey;
        SizeChanged += (_, _) => Redraw();
        BuildMenu();
        Loaded += (_, _) => { Focus(); Select(new CellRef(1, 1)); };
    }

    // ---------- geometry ----------

    /// <summary>Excel stores widths in character units; this is the conversion its own file format documents.</summary>
    private double WidthOf(int column) => Math.Max(MinColumnWidth, Math.Round(_sheet.WidthChars(column) * 7 + 5));

    private IEnumerable<(int Column, double X, double W)> VisibleColumns(double width)
    {
        double x = 0;
        for (int c = _firstColumn; c <= _lastColumn && x < width; c++)
        {
            double w = WidthOf(c);
            yield return (c, x, w);
            x += w;
        }
    }

    /// <summary>
    /// The rows on screen. Frozen rows come first and stay where they are; the rest scroll beneath them. Everything
    /// that draws or hit-tests goes through here, so freezing needs no special case anywhere else.
    /// </summary>
    private IEnumerable<(int Row, double Y)> VisibleRows(double height)
    {
        double y = 0;
        int frozen = FrozenRows;
        for (int r = 1; r <= frozen && y < height; r++, y += RowHeight)
        {
            if (_hidden.Contains(r)) continue;
            yield return (r, y);
        }
        for (int r = Math.Max(_firstRow, frozen + 1); r <= _lastRow && y < height; r++)
        {
            if (_hidden.Contains(r)) continue;
            yield return (r, y);
            y += RowHeight;
        }
    }

    /// <summary>
    /// Rows hidden because they do not match what was filtered for. This is a way of looking at the sheet, not a
    /// change to it: nothing is written to the file, hiding a row never deletes anything, and clearing the filter
    /// puts everything back. A filter that hid rows in the file would be a different and much more dangerous
    /// feature, so this one deliberately is not that.
    /// </summary>
    private readonly HashSet<int> _hidden = new();
    private string _filterText = "";
    private int _filterColumn;

    public bool Filtering => _filterColumn > 0;

    /// <summary>What is being filtered for, to say so in the status bar.</summary>
    public string FilterSaid => Filtering
        ? $"Showing the {CountShown()} rows where column {CellRef.ColumnName(_filterColumn)} has \"{_filterText}\". "
        + $"{_hidden.Count} hidden."
        : "";

    private int CountShown()
    {
        int last = Math.Max(_sheet.Extent.Row, 1);
        return Math.Max(0, last - _hidden.Count);
    }

    /// <summary>
    /// Hide every row whose cell in this column does not contain the text. An empty text clears the filter.
    /// Rows above the first filled row are left alone, so a heading does not vanish.
    /// </summary>
    public void Filter(int column, string text, int firstRow)
    {
        _hidden.Clear();
        _filterColumn = 0;
        _filterText = "";

        if (text.Trim().Length > 0)
        {
            _filterColumn = column;
            _filterText = text.Trim();
            int last = _sheet.Extent.Row;
            for (int r = Math.Max(1, firstRow); r <= last; r++)
            {
                var shown = Get(new CellRef(column, r)).Display;
                if (shown.IndexOf(_filterText, StringComparison.CurrentCultureIgnoreCase) < 0) _hidden.Add(r);
            }
        }

        _firstRow = Math.Max(FrozenRows + 1, 1);
        Reload();
    }

    public void ClearFilter() => Filter(0, "", 1);

    /// <summary>How many rows are held still, never so many that there is no room left to scroll in.</summary>
    internal int FrozenRows
    {
        get
        {
            int want = _sheet.FrozenRows;
            if (want <= 0) return 0;
            int fits = Math.Max(0, (int)(_surface.ActualHeight / RowHeight) - 2);
            return Math.Min(want, fits);
        }
    }

    private Cell Get(CellRef reference)
    {
        if (_cache.TryGetValue(reference, out var cached)) return cached;
        var cell = _sheet.Read(reference);
        _cache[reference] = cell;
        return cell;
    }

    public void Redraw()
    {
        int frozen = FrozenRows;
        _vertical.Minimum = frozen;
        _vertical.Maximum = Math.Max(frozen, _lastRow - Math.Max(1, (int)(_surface.ActualHeight / RowHeight)) + 1);
        _horizontal.Maximum = Math.Max(0, _lastColumn - 6);
        _vertical.ViewportSize = Math.Max(1, _surface.ActualHeight / RowHeight);
        _horizontal.ViewportSize = 6;
        _surface.InvalidateVisual();
        _columns.InvalidateVisual();
        _rows.InvalidateVisual();
        PlaceEditor();
    }

    // ---------- selection and editing ----------

    public void Select(CellRef reference) => Select(reference, extend: false);

    public void Select(CellRef reference, bool extend)
    {
        // A joined block is one cell to a person, so land on the corner that holds the value.
        if (!extend && _sheet.MergeAt(reference) is { } block) reference = block.From;

        _selected = reference;
        if (!extend) _anchor = reference;
        EnsureVisible(reference);
        Redraw();
        SelectionChanged?.Invoke(reference, Get(reference));
    }

    private void EnsureVisible(CellRef reference)
    {
        if (reference.Row < _firstRow) _firstRow = reference.Row;
        int visibleRows = Math.Max(1, (int)(_surface.ActualHeight / RowHeight));
        if (reference.Row >= _firstRow + visibleRows) _firstRow = reference.Row - visibleRows + 1;
        if (reference.Column < _firstColumn) _firstColumn = reference.Column;

        double width = 0;
        int c = _firstColumn;
        while (c < reference.Column && width + WidthOf(c) < _surface.ActualWidth) { width += WidthOf(c); c++; }
        if (c < reference.Column) _firstColumn = reference.Column;

        _vertical.Value = Math.Clamp(_firstRow - 1, 0, _vertical.Maximum);
        _horizontal.Value = Math.Clamp(_firstColumn - 1, 0, _horizontal.Maximum);
    }

    private void SurfaceClick(object sender, MouseButtonEventArgs e)
    {
        Focus();
        var point = e.GetPosition(_surface);
        foreach (var (column, x, w) in VisibleColumns(_surface.ActualWidth))
            if (point.X >= x && point.X < x + w)
                foreach (var (row, y) in VisibleRows(_surface.ActualHeight))
                    if (point.Y >= y && point.Y < y + RowHeight)
                    {
                        Select(new CellRef(column, row), extend: Keyboard.Modifiers == ModifierKeys.Shift);
                        if (e.ClickCount >= 2) BeginEdit(null);
                        return;
                    }
    }

    private void OnKey(object sender, KeyEventArgs e)
    {
        if (_editor.Visibility == Visibility.Visible) return;
        var s = _selected;
        bool shift = (Keyboard.Modifiers & ModifierKeys.Shift) != 0;

        if ((Keyboard.Modifiers & ModifierKeys.Control) != 0)
        {
            switch (e.Key)
            {
                case Key.Up: Jump(0, -1, shift); e.Handled = true; return;
                case Key.Down: Jump(0, 1, shift); e.Handled = true; return;
                case Key.Left: Jump(-1, 0, shift); e.Handled = true; return;
                case Key.Right: Jump(1, 0, shift); e.Handled = true; return;
                case Key.Home: Select(new CellRef(1, 1), shift); e.Handled = true; return;
                case Key.End:
                {
                    var end = _sheet.Extent;
                    Select(new CellRef(Math.Max(1, end.Column), Math.Max(1, end.Row)), shift);
                    e.Handled = true; return;
                }
                case Key.C: Copy(); e.Handled = true; return;
                case Key.X: Copy(); ClearRange(); e.Handled = true; return;
                case Key.V: Paste(); e.Handled = true; return;
                case Key.A:
                    var extent = _sheet.Extent;
                    _anchor = new CellRef(1, 1);
                    Select(new CellRef(Math.Max(1, extent.Column), Math.Max(1, extent.Row)), extend: true);
                    e.Handled = true; return;
            }
            return;
        }

        switch (e.Key)
        {
            case Key.Up: Select(new CellRef(s.Column, Math.Max(1, s.Row - 1)), shift); e.Handled = true; return;
            case Key.Down or Key.Enter: Select(new CellRef(s.Column, Math.Min(_lastRow, s.Row + 1)), shift); e.Handled = true; return;
            case Key.Left: Select(new CellRef(Math.Max(1, s.Column - 1), s.Row), shift); e.Handled = true; return;
            case Key.Right or Key.Tab: Select(new CellRef(Math.Min(_lastColumn, s.Column + 1), s.Row), shift); e.Handled = true; return;
            case Key.Home: Select(new CellRef(1, s.Row), shift); e.Handled = true; return;
            case Key.F2: BeginEdit(null); e.Handled = true; return;
            case Key.Delete or Key.Back: ClearRange(); e.Handled = true; return;
            case Key.PageDown: Select(new CellRef(s.Column, Math.Min(_lastRow, s.Row + 20)), shift); e.Handled = true; return;
            case Key.PageUp: Select(new CellRef(s.Column, Math.Max(1, s.Row - 20)), shift); e.Handled = true; return;
        }
    }

    // ---------- copy and paste ----------

    /// <summary>
    /// Copy the selection as tab separated text, which is what every spreadsheet reads and writes on the clipboard.
    /// A cell with a formula copies as its formula, the way Excel copies one, so pasting it back puts a formula back.
    /// </summary>
    /// <summary>Every cell in the selection, which is one cell unless a block is picked out.</summary>
    public IEnumerable<CellRef> SelectedCells()
    {
        var (left, top, right, bottom) = Range;
        for (int r = top; r <= bottom; r++)
            for (int c = left; c <= right; c++)
                yield return new CellRef(c, r);
    }

    /// <summary>
    /// Put whatever is in the top cell of the selection into the rest of it. Formulas move with it the way they do in
    /// Excel, so a rate filled down a column reads the row it lands on rather than the row it came from.
    /// </summary>
    public int FillDown()
    {
        var (left, top, right, bottom) = Range;
        if (bottom == top) return 0;

        var changes = new List<(CellRef Cell, string Value)>();
        for (int c = left; c <= right; c++)
        {
            var from = new CellRef(c, top);
            var source = _sheet.Read(from);
            string typed = source.Formula ?? source.Raw;
            if (typed.Length == 0) continue;

            for (int r = top + 1; r <= bottom; r++)
            {
                string value = source.Formula is null
                    ? typed
                    : Refs.Rewrite(typed, token =>
                      {
                          // A relative row follows the fill; a row pinned with a dollar stays where it was.
                          if (token.Kind != Refs.Shape.Cell) return null;
                          int shift = r - top;
                          var range = token.Range;
                          int min = token.Row1Fixed ? range.RowMin : range.RowMin + shift;
                          int max = token.Row2Fixed ? range.RowMax : range.RowMax + shift;
                          if (min == range.RowMin && max == range.RowMax) return null;
                          if (min < 1 || max > 1048576) return "#REF!";
                          return Refs.Write(range with { RowMin = min, RowMax = max }, token.Kind,
                                            token.Column1Fixed, token.Row1Fixed, token.Column2Fixed, token.Row2Fixed,
                                            token.IsRange);
                      });
                changes.Add((new CellRef(c, r), value));
            }
        }

        if (changes.Count == 0) return 0;
        ApplyMany(changes);
        return changes.Count;
    }

    /// <summary>Read the sheet again from the file, after something changed it from outside this view.</summary>
    public void Reload()
    {
        _cache.Clear();
        var extent = _sheet.Extent;
        _lastColumn = Math.Max(extent.Column + 6, 12);
        _lastRow = Math.Max(extent.Row + 40, 60);
        Redraw();
        SelectionChanged?.Invoke(_selected, Get(_selected));
    }

    /// <summary>Put a row or column in, or take one out, where the selection is.</summary>
    public event Action<GridEdit, int>? GridChangeRequested;

    /// <summary>Asked to filter on a column; the window asks for the text, because it owns the dialogs.</summary>
    public event Action<int, int>? FilterRequested;

    /// <summary>Asked to sort the selection by a column; the window does it, because it owns undo and the message.</summary>
    public event Action<int, int, int, int, int, bool>? SortRequested;

    private void BuildMenu()
    {
        var menu = new ContextMenu();
        void Item(string text, GridEdit edit, Func<int> where)
        {
            var entry = new MenuItem { Header = text };
            entry.Click += (_, _) => GridChangeRequested?.Invoke(edit, where());
            menu.Items.Add(entry);
        }
        Item("Insert row above", GridEdit.InsertRow, () => Range.Top);
        Item("Insert row below", GridEdit.InsertRow, () => Range.Bottom + 1);
        Item("Delete this row", GridEdit.DeleteRow, () => Range.Top);
        menu.Items.Add(new Separator());
        Item("Insert column left", GridEdit.InsertColumn, () => Range.Left);
        Item("Insert column right", GridEdit.InsertColumn, () => Range.Right + 1);
        Item("Delete this column", GridEdit.DeleteColumn, () => Range.Left);
        menu.Items.Add(new Separator());

        var fit = new MenuItem { Header = "Fit this column to its contents" };
        fit.Click += (_, _) =>
        {
            int column = Range.Left;
            double before = _sheet.WidthChars(column);
            _sheet.SetWidthChars(column, _sheet.WidestChars(column));
            NoteWidthChange(column, before);
        };
        menu.Items.Add(fit);

        void SortItem(string text, bool ascending)
        {
            var entry = new MenuItem { Header = text };
            entry.Click += (_, _) =>
            {
                var (l, t, r, b) = Range;
                // Sorting one column on its own would tear the row apart, so a single-column selection sorts the
                // whole width of what is filled in, keeping each row together.
                if (l == r) { l = 1; r = Math.Max(1, _sheet.Extent.Column); }
                SortRequested?.Invoke(t, b, l, r, Range.Left, ascending);
            };
            menu.Items.Add(entry);
        }
        menu.Items.Add(new Separator());
        var filter = new MenuItem { Header = "Show only rows where this column..." };
        filter.Click += (_, _) => FilterRequested?.Invoke(Range.Left, Range.Top);
        menu.Items.Add(filter);
        menu.Opened += (_, _) =>
            filter.Header = Filtering ? "Show every row again" : "Show only rows where this column...";
        menu.Items.Add(new Separator());
        SortItem("Sort these rows by this column", true);
        SortItem("Sort these rows by this column, backwards", false);
        menu.Items.Add(new Separator());

        // How the cells look, and how tall and wide things are.
        var look = new MenuItem { Header = "How these cells look" };
        void Look(string text, Action what)
        {
            var entry = new MenuItem { Header = text };
            entry.Click += (_, _) => what();
            look.Items.Add(entry);
        }
        void Align(string text, string where)
        {
            Look(text, () =>
            {
                var cells = InRange().ToList();
                var was = cells.Select(c => _sheet.AlignmentAt(c)).ToList();
                _sheet.SetAlignment(cells, where, null);
                AfterLook(cells, was);
            });
        }
        Align("Left", "left");
        Align("Centred", "center");
        Align("Right", "right");
        Align("However its style says", "general");
        look.Items.Add(new Separator());
        Look("Wrap the words onto more lines", () =>
        {
            var cells = InRange().ToList();
            var was = cells.Select(c => _sheet.AlignmentAt(c)).ToList();
            bool on = !_sheet.AlignmentAt(_selected).Wrap;
            _sheet.SetAlignment(cells, null, on);
            AfterLook(cells, was);
        });
        look.Items.Add(new Separator());
        foreach (var (name, fill) in new[]
                 { ("Yellow", "FFF3C4"), ("Green", "D9EAD3"), ("Blue", "DCE7F5"), ("Pink", "F7DDE3"), ("Grey", "EDEDED") })
        {
            var colour = fill;
            Look(name + " behind them", () =>
            {
                var cells = InRange().ToList();
                var was = cells.Select(c => _sheet.ColoursAt(c)).ToList();
                _sheet.SetColours(cells, colour, null);
                AfterColour(cells, was);
            });
        }
        Look("No colour behind them", () =>
        {
            var cells = InRange().ToList();
            var was = cells.Select(c => _sheet.ColoursAt(c)).ToList();
            _sheet.SetColours(cells, "", null);
            AfterColour(cells, was);
        });
        menu.Items.Add(look);

        var freezeColumns = new MenuItem { Header = "Keep the columns left of this one on screen" };
        freezeColumns.Click += (_, _) =>
        {
            int before = _sheet.FrozenColumns;
            int want = before > 0 ? 0 : Math.Max(0, Range.Left - 1);
            int rows = _sheet.FrozenRows;
            _sheet.SetFrozen(want, rows);
            _firstColumn = Math.Max(want + 1, _firstColumn);
            Redraw();
            Raise(() => { _sheet.SetFrozen(before, rows); Redraw(); },
                  () => { _sheet.SetFrozen(want, rows); Redraw(); });
        };
        menu.Items.Add(freezeColumns);
        menu.Opened += (_, _) => freezeColumns.Header = _sheet.FrozenColumns > 0
            ? "Let every column scroll again" : "Keep the columns left of this one on screen";

        var freeze = new MenuItem { Header = "Keep the rows above this one on screen" };
        freeze.Click += (_, _) =>
        {
            int before = _sheet.FrozenRows;
            // Freezing "above this row" means the rows before the selected one stay put, which is what people mean
            // when they click on row 2 and ask for the header to stay.
            int want = before > 0 ? 0 : Math.Max(0, Range.Top - 1);
            _sheet.SetFrozenRows(want);
            _firstRow = Math.Max(want + 1, _firstRow);
            Redraw();
            Edited?.Invoke(new Edit(
                () => { _sheet.SetFrozenRows(before); Redraw(); },
                () => { _sheet.SetFrozenRows(want); Redraw(); }));
        };
        menu.Items.Add(freeze);
        menu.Opened += (_, _) =>
            freeze.Header = _sheet.FrozenRows > 0 ? "Let every row scroll again" : "Keep the rows above this one on screen";

        ContextMenu = menu;
    }

    public void Copy()
    {
        var (left, top, right, bottom) = Range;
        var rows = new List<IReadOnlyList<string>>();
        for (int r = top; r <= bottom; r++)
        {
            var line = new List<string>();
            for (int c = left; c <= right; c++)
            {
                var cell = Get(new CellRef(c, r));
                line.Add(cell.Formula ?? cell.Display);
            }
            rows.Add(line);
        }
        try { Clipboard.SetText(Tabular.Write(rows)); } catch { /* another program may hold the clipboard */ }
    }

    /// <summary>Paste tab separated text, filling cells from the top left of the selection. One step of undo.</summary>
    public void Paste()
    {
        string text;
        try { text = Clipboard.ContainsText() ? Clipboard.GetText() : ""; } catch { return; }
        if (text.Length == 0) return;

        var (left, top, _, _) = Range;
        var rows = Tabular.Read(text);
        var changes = new List<(CellRef Cell, string Value)>();
        for (int r = 0; r < rows.Count; r++)
            for (int c = 0; c < rows[r].Count; c++)
            {
                int column = left + c, row = top + r;
                if (column > 16384 || row > 1048576) continue;
                changes.Add((new CellRef(column, row), rows[r][c]));
            }
        if (changes.Count == 0) return;

        ApplyMany(changes);
        _anchor = new CellRef(left, top);
        Select(new CellRef(Math.Min(16384, left + Tabular.Width(rows) - 1),
                           Math.Min(1048576, top + rows.Count - 1)), extend: true);
    }

    /// <summary>Empty every cell in the selection, in one step of undo.</summary>
    public void ClearRange()
    {
        var (left, top, right, bottom) = Range;
        var changes = new List<(CellRef, string)>();
        for (int r = top; r <= bottom; r++)
            for (int c = left; c <= right; c++)
                if (!Get(new CellRef(c, r)).IsEmpty || Get(new CellRef(c, r)).Formula is not null)
                    changes.Add((new CellRef(c, r), ""));
        if (changes.Count > 0) ApplyMany(changes);
    }

    protected override void OnTextInput(TextCompositionEventArgs e)
    {
        base.OnTextInput(e);
        if (_editor.Visibility == Visibility.Visible) return;
        if (e.Text.Length > 0 && !char.IsControl(e.Text[0])) { BeginEdit(e.Text); e.Handled = true; }
    }

    /// <summary>Start typing into the selected cell. A seed character replaces the contents, F2 keeps them.</summary>
    public void BeginEdit(string? seed)
    {
        var cell = Get(_selected);
        _editor.Text = seed ?? cell.Formula ?? cell.Raw;
        _editor.Visibility = Visibility.Visible;
        PlaceEditor();
        _editor.Focus();
        _editor.CaretIndex = _editor.Text.Length;
    }

    private void PlaceEditor()
    {
        if (_editor.Visibility != Visibility.Visible) return;
        double x = 0;
        bool found = false;
        foreach (var (column, cx, w) in VisibleColumns(_surface.ActualWidth))
            if (column == _selected.Column) { x = cx; _editor.Width = Math.Max(w, 60); found = true; break; }
        double y = 0;
        bool foundRow = false;
        foreach (var (row, ry) in VisibleRows(_surface.ActualHeight))
            if (row == _selected.Row) { y = ry; foundRow = true; break; }
        if (!found || !foundRow) { CancelEdit(); return; }
        _editor.Margin = new Thickness(x - 1, y - 1, 0, 0);
    }

    private void EditorKey(object sender, KeyEventArgs e)
    {
        if (e.Key == Key.Escape) { CancelEdit(); e.Handled = true; }
        else if (e.Key == Key.Enter)
        {
            var next = new CellRef(_selected.Column, Math.Min(_lastRow, _selected.Row + 1));
            CommitEditor();
            Select(next);
            Focus();
            e.Handled = true;
        }
        else if (e.Key == Key.Tab)
        {
            var next = new CellRef(Math.Min(_lastColumn, _selected.Column + 1), _selected.Row);
            CommitEditor();
            Select(next);
            Focus();
            e.Handled = true;
        }
    }

    /// <summary>
    /// Move the way Ctrl with an arrow moves in a spreadsheet: to the far end of the run of filled cells you are
    /// in, or, if the next cell is empty, across the gap to the next thing there is. It is how you get to the
    /// bottom of ten thousand rows without holding a key down.
    /// </summary>
    private void Jump(int dx, int dy, bool extend)
    {
        int column = _selected.Column, row = _selected.Row;
        bool Filled(int c, int r) => c >= 1 && r >= 1 && c <= 16384 && r <= 1048576
                                     && Get(new CellRef(c, r)).Kind != CellKind.Empty;

        int lastColumn = column, lastRow = row;
        bool startedFilled = Filled(column, row);
        // A run of filled cells ends at the last filled one; a gap ends at the first filled one.
        for (int steps = 0; steps < 200_000; steps++)
        {
            int nextColumn = column + dx, nextRow = row + dy;
            if (nextColumn < 1 || nextRow < 1 || nextColumn > 16384 || nextRow > 1048576) break;

            bool next = Filled(nextColumn, nextRow);
            column = nextColumn; row = nextRow;
            if (startedFilled) { if (!next) { column -= dx; row -= dy; break; } lastColumn = column; lastRow = row; }
            else if (next) break;
        }

        // Never wander past where the sheet has anything, or a stray keystroke lands you at row a million.
        var extent = _sheet.Extent;
        column = Math.Clamp(column, 1, Math.Max(1, Math.Max(extent.Column, _lastColumn)));
        row = Math.Clamp(row, 1, Math.Max(1, Math.Max(extent.Row, _lastRow)));
        Select(new CellRef(column, row), extend);
    }

    /// <summary>Go to a cell by name, for Ctrl+G. False when that is not a cell reference.</summary>
    public bool GoTo(string reference)
    {
        if (!CellRef.TryParse(reference.Trim().Replace("$", ""), out var cell)) return false;
        Select(cell);
        Focus();
        return true;
    }

    private void CancelEdit() { _editor.Visibility = Visibility.Collapsed; Focus(); }

    private void CommitEditor()
    {
        if (_editor.Visibility != Visibility.Visible) return;
        string typed = _editor.Text;
        _editor.Visibility = Visibility.Collapsed;
        Apply(typed);
    }

    /// <summary>Put a typed value into the selected cell and redraw the neighbourhood.</summary>
    public void Apply(string typed) => ApplyMany(new[] { (_selected, typed) });

    /// <summary>
    /// Change a set of cells together. A paste is one action to the person who did it, so it is one step of undo
    /// rather than one per cell.
    /// </summary>
    public void ApplyMany(IReadOnlyList<(CellRef Cell, string Value)> changes)
    {
        var before = new List<(CellRef Cell, string Value)>();
        foreach (var (cell, value) in changes)
        {
            var current = Get(cell);
            string was = current.Formula ?? current.Raw;
            if (was == value) continue;
            before.Add((cell, was));
        }
        if (before.Count == 0) return;

        foreach (var (cell, value) in changes) _sheet.Set(cell, value);
        AfterChange();

        var undoTo = before;
        var redoTo = changes.ToList();
        var landOn = _selected;
        Edited?.Invoke(new Edit(
            () =>
            {
                foreach (var (cell, was) in undoTo) _sheet.Set(cell, was);
                AfterChange();
                Select(landOn);
            },
            () =>
            {
                foreach (var (cell, value) in redoTo) _sheet.Set(cell, value);
                AfterChange();
                Select(landOn);
            }));
    }

    private void AfterChange()
    {
        _cache.Clear();   // a formula elsewhere may now show differently
        var extent = _sheet.Extent;
        _lastColumn = Math.Max(_lastColumn, extent.Column + 6);
        _lastRow = Math.Max(_lastRow, extent.Row + 40);
        Redraw();
        SelectionChanged?.Invoke(_selected, Get(_selected));
    }

    // ---------- drawing ----------

    private static readonly Typeface Face = new(new FontFamily("Segoe UI Variable Text, Segoe UI"),
        FontStyles.Normal, FontWeights.Normal, FontStretches.Normal);
    private static readonly Typeface BoldFace = new(new FontFamily("Segoe UI Variable Text, Segoe UI"),
        FontStyles.Normal, FontWeights.SemiBold, FontStretches.Normal);

    private FormattedText Text(string text, Brush brush, bool bold = false) =>
        new(text, CultureInfo.CurrentCulture, FlowDirection.LeftToRight, bold ? BoldFace : Face, 12.5, brush,
            VisualTreeHelper.GetDpi(this).PixelsPerDip);

    /// <summary>Six hex digits as a brush, or nothing when they are not six hex digits.</summary>
    private static SolidColorBrush? Colour(string hex)
    {
        if (hex.Length != 6 || !hex.All(Uri.IsHexDigit)) return null;
        try
        {
            return new SolidColorBrush(Color.FromRgb(
                Convert.ToByte(hex[..2], 16), Convert.ToByte(hex[2..4], 16), Convert.ToByte(hex[4..], 16)));
        }
        catch { return null; }
    }

    private sealed class Surface : FrameworkElement
    {
        private readonly SheetView _view;
        public Surface(SheetView view) { _view = view; ClipToBounds = true; }

        protected override void OnRender(DrawingContext dc)
        {
            var v = _view;
            double w = ActualWidth, h = ActualHeight;
            dc.DrawRectangle(App.B("Surface"), null, new Rect(0, 0, w, h));
            var line = new Pen(App.B("Line"), 1);
            var selection = new Pen(App.B("Accent"), 2);

            foreach (var (row, y) in v.VisibleRows(h))
                dc.DrawLine(line, new Point(0, Snap(y + RowHeight)), new Point(w, Snap(y + RowHeight)));

            foreach (var (column, x, cw) in v.VisibleColumns(w))
            {
                dc.DrawLine(line, new Point(Snap(x + cw), 0), new Point(Snap(x + cw), h));

                foreach (var (row, y) in v.VisibleRows(h))
                {
                    var here = new CellRef(column, row);
                    // Joined blocks are drawn after the gridlines, so no line runs through the middle of one.
                    if (v._sheet.MergeAt(here) is not null) continue;

                    var cell = v.Get(here);

                    // A colour behind the cell is drawn whether or not there is anything in it, because an empty
                    // cell someone coloured is still coloured.
                    var (fill, inkColour) = v._sheet.ColoursAt(here);
                    if (fill.Length == 6 && Colour(fill) is { } paint)
                        dc.DrawRectangle(paint, null, new Rect(Snap(x) + 1, Snap(y) + 1, cw - 1, RowHeight - 1));

                    if (cell.Display.Length == 0) continue;

                    bool number = cell.Kind is CellKind.Number or CellKind.Formula or CellKind.Boolean
                                  && double.TryParse(cell.Raw, NumberStyles.Float, CultureInfo.InvariantCulture, out _);
                    // A formula with no computed value shows as the formula, in the quieter ink, so it reads as
                    // "this is worked out when you open it in Excel" rather than as a value.
                    bool uncomputed = cell.Kind == CellKind.Formula && cell.Raw.Length == 0;
                    var brush = cell.Kind == CellKind.Error ? Brushes.IndianRed
                              : uncomputed ? App.B("Ink3")
                              : inkColour.Length == 6 && Colour(inkColour) is { } chosen ? chosen
                              : App.B("Ink");
                    double drawWidth = cw;
                    var text = v.Text(cell.Display, brush);
                    text.MaxTextWidth = Math.Max(4, drawWidth - 10);
                    text.MaxTextHeight = RowHeight;
                    text.Trimming = TextTrimming.CharacterEllipsis;
                    // What the cell says about where it sits wins; a number falls back to the right, as it does
                    // in every spreadsheet, and text to the left.
                    var (where, wraps) = v._sheet.AlignmentAt(here);
                    if (wraps) { text.MaxTextHeight = RowHeight; text.Trimming = TextTrimming.CharacterEllipsis; }
                    double tx = where switch
                    {
                        "center" => x + (drawWidth - text.Width) / 2,
                        "right" => x + drawWidth - 5 - text.Width,
                        "left" => x + 5,
                        _ => number ? x + drawWidth - 5 - text.Width : x + 5,
                    };
                    dc.DrawText(text, new Point(Math.Max(x + 5, tx), y + 4));
                }
            }

            // A stronger line under the rows that are held still, so it reads as a header that stays rather than
            // as a sheet that has lost its place.
            int frozenHere = v.FrozenRows;
            if (frozenHere > 0)
            {
                double edge = Snap(frozenHere * RowHeight);
                dc.DrawLine(new Pen(App.B("Ink3"), 1.5), new Point(0, edge), new Point(w, edge));
            }

            // Joined blocks: paint over the grid, put the border back round the outside, and write the value once
            // across the whole width. Excel shows a merged heading as one cell and so should this.
            foreach (var (from, to) in v._sheet.Merges)
            {
                double x1 = double.NaN, x2 = double.NaN, y1 = double.NaN, y2 = double.NaN;
                foreach (var (column, cx, cw) in v.VisibleColumns(w))
                {
                    if (column == from.Column) x1 = cx;
                    if (column == to.Column) x2 = cx + cw;
                }
                foreach (var (row, cy) in v.VisibleRows(h))
                {
                    if (row == from.Row) y1 = cy;
                    if (row == to.Row) y2 = cy + RowHeight;
                }
                if (double.IsNaN(x1) || double.IsNaN(x2) || double.IsNaN(y1) || double.IsNaN(y2)) continue;

                var area = new Rect(Snap(x1), Snap(y1), Math.Max(1, x2 - x1), Math.Max(1, y2 - y1));
                dc.DrawRectangle(App.B("Surface"), null, area);
                dc.DrawLine(line, new Point(area.Right, area.Top), new Point(area.Right, area.Bottom));
                dc.DrawLine(line, new Point(area.Left, area.Bottom), new Point(area.Right, area.Bottom));

                var held = v.Get(from);
                if (held.Display.Length == 0) continue;
                var joinedText = v.Text(held.Display, App.B("Ink"));
                joinedText.MaxTextWidth = Math.Max(4, area.Width - 10);
                joinedText.MaxTextHeight = area.Height;
                joinedText.Trimming = TextTrimming.CharacterEllipsis;
                dc.DrawText(joinedText, new Point(area.Left + 5, area.Top + 4));
            }

            // The selection sits on top so its outline is never cut by a gridline. A range is tinted; the cell you
            // are actually on keeps the crisp outline so you can always see where typing would go.
            var (left, top, right, bottom) = v.Range;
            if (v.HasRange)
            {
                var tint = new SolidColorBrush(((SolidColorBrush)App.B("Accent")).Color) { Opacity = 0.12 };
                foreach (var (column, x, cw) in v.VisibleColumns(w))
                    if (column >= left && column <= right)
                        foreach (var (row, y) in v.VisibleRows(h))
                            if (row >= top && row <= bottom)
                                dc.DrawRectangle(tint, null, new Rect(Snap(x) + 1, Snap(y) + 1, cw - 2, RowHeight - 2));
            }

            foreach (var (column, x, cw) in v.VisibleColumns(w))
                if (column == v._selected.Column)
                    foreach (var (row, y) in v.VisibleRows(h))
                        if (row == v._selected.Row)
                            dc.DrawRectangle(null, selection, new Rect(Snap(x) + 1, Snap(y) + 1, cw - 2, RowHeight - 2));
        }

        private static double Snap(double value) => Math.Round(value) + 0.5;
    }

    /// <summary>Every cell in the selection, which is what the look-changing menu items work on.</summary>
    private IEnumerable<CellRef> InRange()
    {
        var (left, top, right, bottom) = Range;
        for (int r = top; r <= bottom; r++)
            for (int c = left; c <= right; c++)
                yield return new CellRef(c, r);
    }

    /// <summary>Put alignment back the way it was, cell by cell, which is what undo needs.</summary>
    private void AfterLook(List<CellRef> cells, List<(string Horizontal, bool Wrap)> was)
    {
        var now = cells.Select(c => _sheet.AlignmentAt(c)).ToList();
        Redraw();
        Raise(() =>
        {
            for (int i = 0; i < cells.Count; i++) _sheet.SetAlignment(new[] { cells[i] }, was[i].Horizontal, was[i].Wrap);
            Redraw();
        },
        () =>
        {
            for (int i = 0; i < cells.Count; i++) _sheet.SetAlignment(new[] { cells[i] }, now[i].Horizontal, now[i].Wrap);
            Redraw();
        });
    }

    private void AfterColour(List<CellRef> cells, List<(string Background, string Ink)> was)
    {
        var now = cells.Select(c => _sheet.ColoursAt(c)).ToList();
        Redraw();
        Raise(() =>
        {
            for (int i = 0; i < cells.Count; i++) _sheet.SetColours(new[] { cells[i] }, was[i].Background, was[i].Ink);
            Redraw();
        },
        () =>
        {
            for (int i = 0; i < cells.Count; i++) _sheet.SetColours(new[] { cells[i] }, now[i].Background, now[i].Ink);
            Redraw();
        });
    }

    /// <summary>A width change is an edit like any other: it dirties the file and Ctrl+Z puts the old width back.</summary>
    internal void NoteWidthChange(int column, double before)
    {
        double now = _sheet.WidthChars(column);
        Redraw();
        Edited?.Invoke(new Edit(
            () => { _sheet.SetWidthChars(column, before); Redraw(); },
            () => { _sheet.SetWidthChars(column, now); Redraw(); }));
    }

    private sealed class Strip : FrameworkElement
    {
        private readonly SheetView _view;
        private readonly bool _horizontal;
        public Strip(SheetView view, bool horizontal)
        {
            _view = view; _horizontal = horizontal; ClipToBounds = true;
            if (!horizontal) return;

            // Dragging the line between two column headings changes the width, and double-clicking it fits the
            // column to its longest value. The file keeps widths, so this is an edit like any other: it marks the
            // sheet changed and is written back when you save.
            MouseMove += (_, e) =>
            {
                if (_dragging > 0)
                {
                    double want = (e.GetPosition(this).X - _dragFrom) / 7.0;
                    _view._sheet.SetWidthChars(_dragging, Math.Max(0.5, want));
                    _view.Redraw();
                    return;
                }
                Cursor = EdgeAt(e.GetPosition(this).X) > 0 ? Cursors.SizeWE : Cursors.Arrow;
            };
            MouseLeftButtonDown += (_, e) =>
            {
                int column = EdgeAt(e.GetPosition(this).X);
                if (column == 0) return;
                if (e.ClickCount >= 2)
                {
                    double before = _view._sheet.WidthChars(column);
                    _view._sheet.SetWidthChars(column, _view._sheet.WidestChars(column));
                    _view.NoteWidthChange(column, before);
                    e.Handled = true;
                    return;
                }
                _dragging = column;
                _widthBefore = _view._sheet.WidthChars(column);
                foreach (var (c, x, _) in _view.VisibleColumns(ActualWidth)) if (c == column) _dragFrom = x;
                CaptureMouse();
                e.Handled = true;
            };
            MouseLeftButtonUp += (_, e) =>
            {
                if (_dragging == 0) return;
                int done = _dragging;
                _dragging = 0;
                ReleaseMouseCapture();
                _view.NoteWidthChange(done, _widthBefore);
                e.Handled = true;
            };
        }

        private int _dragging;
        private double _dragFrom;
        private double _widthBefore;

        /// <summary>Which column's right edge is under this x, within a few pixels either side. Nought for none.</summary>
        private int EdgeAt(double x)
        {
            foreach (var (column, cx, cw) in _view.VisibleColumns(ActualWidth))
                if (Math.Abs(x - (cx + cw)) <= 3) return column;
            return 0;
        }

        protected override void OnRender(DrawingContext dc)
        {
            var v = _view;
            double w = ActualWidth, h = ActualHeight;
            dc.DrawRectangle(App.B("Surface2"), null, new Rect(0, 0, w, h));
            var line = new Pen(App.B("Line"), 1);

            if (_horizontal)
            {
                foreach (var (column, x, cw) in v.VisibleColumns(w))
                {
                    dc.DrawLine(line, new Point(Snap(x + cw), 0), new Point(Snap(x + cw), h));
                    bool on = column >= v.Range.Left && column <= v.Range.Right;
                    var t = v.Text(CellRef.ColumnName(column), on ? App.B("AccentInk") : App.B("Ink2"), on);
                    dc.DrawText(t, new Point(x + (cw - t.Width) / 2, (h - t.Height) / 2));
                }
                dc.DrawLine(line, new Point(0, Snap(h)), new Point(w, Snap(h)));
            }
            else
            {
                foreach (var (row, y) in v.VisibleRows(h))
                {
                    dc.DrawLine(line, new Point(0, Snap(y + RowHeight)), new Point(w, Snap(y + RowHeight)));
                    bool on = row >= v.Range.Top && row <= v.Range.Bottom;
                    var t = v.Text(row.ToString(CultureInfo.CurrentCulture), on ? App.B("AccentInk") : App.B("Ink2"), on);
                    dc.DrawText(t, new Point(w - 7 - t.Width, y + 4));
                }
                dc.DrawLine(line, new Point(Snap(w), 0), new Point(Snap(w), h));
            }
        }

        private static double Snap(double value) => Math.Round(value) + 0.5;
    }
}
