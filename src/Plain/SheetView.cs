using System.Globalization;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Input;
using System.Windows.Media;
using Plain.Core;

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
    private int _firstColumn = 1, _firstRow = 1;
    private int _lastColumn = 40, _lastRow = 200;

    public event Action<CellRef, Cell>? SelectionChanged;
    public event Action<Action>? Edited;

    public Sheet Sheet => _sheet;
    public CellRef Selected => _selected;

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
        _vertical.Scroll += (_, _) => { _firstRow = (int)_vertical.Value + 1; Redraw(); };
        _horizontal.Scroll += (_, _) => { _firstColumn = (int)_horizontal.Value + 1; Redraw(); };

        Background = App.B("Surface");
        Focusable = true;
        _surface.MouseLeftButtonDown += SurfaceClick;
        _surface.MouseWheel += (_, e) => { _vertical.Value = Math.Clamp(_vertical.Value - e.Delta / 40.0, 0, _vertical.Maximum); _firstRow = (int)_vertical.Value + 1; Redraw(); };
        KeyDown += OnKey;
        SizeChanged += (_, _) => Redraw();
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

    private IEnumerable<(int Row, double Y)> VisibleRows(double height)
    {
        double y = 0;
        for (int r = _firstRow; r <= _lastRow && y < height; r++, y += RowHeight) yield return (r, y);
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
        _vertical.Maximum = Math.Max(0, _lastRow - Math.Max(1, (int)(_surface.ActualHeight / RowHeight)) + 1);
        _horizontal.Maximum = Math.Max(0, _lastColumn - 6);
        _vertical.ViewportSize = Math.Max(1, _surface.ActualHeight / RowHeight);
        _horizontal.ViewportSize = 6;
        _surface.InvalidateVisual();
        _columns.InvalidateVisual();
        _rows.InvalidateVisual();
        PlaceEditor();
    }

    // ---------- selection and editing ----------

    public void Select(CellRef reference)
    {
        _selected = reference;
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
                        Select(new CellRef(column, row));
                        if (e.ClickCount >= 2) BeginEdit(null);
                        return;
                    }
    }

    private void OnKey(object sender, KeyEventArgs e)
    {
        if (_editor.Visibility == Visibility.Visible) return;
        var s = _selected;
        switch (e.Key)
        {
            case Key.Up: Select(new CellRef(s.Column, Math.Max(1, s.Row - 1))); e.Handled = true; return;
            case Key.Down or Key.Enter: Select(new CellRef(s.Column, Math.Min(_lastRow, s.Row + 1))); e.Handled = true; return;
            case Key.Left: Select(new CellRef(Math.Max(1, s.Column - 1), s.Row)); e.Handled = true; return;
            case Key.Right or Key.Tab: Select(new CellRef(Math.Min(_lastColumn, s.Column + 1), s.Row)); e.Handled = true; return;
            case Key.Home: Select(new CellRef(1, s.Row)); e.Handled = true; return;
            case Key.F2: BeginEdit(null); e.Handled = true; return;
            case Key.Delete or Key.Back: Apply(""); e.Handled = true; return;
            case Key.PageDown: Select(new CellRef(s.Column, Math.Min(_lastRow, s.Row + 20))); e.Handled = true; return;
            case Key.PageUp: Select(new CellRef(s.Column, Math.Max(1, s.Row - 20))); e.Handled = true; return;
        }
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

    private void CancelEdit() { _editor.Visibility = Visibility.Collapsed; Focus(); }

    private void CommitEditor()
    {
        if (_editor.Visibility != Visibility.Visible) return;
        string typed = _editor.Text;
        _editor.Visibility = Visibility.Collapsed;
        Apply(typed);
    }

    /// <summary>Put a typed value into the selected cell and redraw the neighbourhood.</summary>
    public void Apply(string typed)
    {
        var current = Get(_selected);
        string was = current.Formula ?? current.Raw;
        if (was == typed) return;
        var where = _selected;
        _sheet.Set(_selected, typed);
        _cache.Clear();   // a formula elsewhere may now show differently
        var extent = _sheet.Extent;
        _lastColumn = Math.Max(_lastColumn, extent.Column + 6);
        _lastRow = Math.Max(_lastRow, extent.Row + 40);
        Redraw();
        SelectionChanged?.Invoke(_selected, Get(_selected));
        Edited?.Invoke(() =>
        {
            _sheet.Set(where, was);
            _cache.Clear();
            Select(where);
            Redraw();
        });
    }

    // ---------- drawing ----------

    private static readonly Typeface Face = new(new FontFamily("Segoe UI Variable Text, Segoe UI"),
        FontStyles.Normal, FontWeights.Normal, FontStretches.Normal);
    private static readonly Typeface BoldFace = new(new FontFamily("Segoe UI Variable Text, Segoe UI"),
        FontStyles.Normal, FontWeights.SemiBold, FontStretches.Normal);

    private FormattedText Text(string text, Brush brush, bool bold = false) =>
        new(text, CultureInfo.CurrentCulture, FlowDirection.LeftToRight, bold ? BoldFace : Face, 12.5, brush,
            VisualTreeHelper.GetDpi(this).PixelsPerDip);

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
                    var cell = v.Get(new CellRef(column, row));
                    if (cell.Display.Length == 0) continue;

                    bool number = cell.Kind is CellKind.Number or CellKind.Formula or CellKind.Boolean
                                  && double.TryParse(cell.Raw, NumberStyles.Float, CultureInfo.InvariantCulture, out _);
                    // A formula with no computed value shows as the formula, in the quieter ink, so it reads as
                    // "this is worked out when you open it in Excel" rather than as a value.
                    bool uncomputed = cell.Kind == CellKind.Formula && cell.Raw.Length == 0;
                    var brush = cell.Kind == CellKind.Error ? Brushes.IndianRed
                              : uncomputed ? App.B("Ink3")
                              : App.B("Ink");
                    var text = v.Text(cell.Display, brush);
                    text.MaxTextWidth = Math.Max(4, cw - 10);
                    text.MaxTextHeight = RowHeight;
                    text.Trimming = TextTrimming.CharacterEllipsis;
                    double tx = number ? x + cw - 5 - text.Width : x + 5;
                    dc.DrawText(text, new Point(Math.Max(x + 5, tx), y + 4));
                }
            }

            // The selected cell sits on top so its outline is never cut by a gridline.
            foreach (var (column, x, cw) in v.VisibleColumns(w))
                if (column == v._selected.Column)
                    foreach (var (row, y) in v.VisibleRows(h))
                        if (row == v._selected.Row)
                            dc.DrawRectangle(null, selection, new Rect(Snap(x) + 1, Snap(y) + 1, cw - 2, RowHeight - 2));
        }

        private static double Snap(double value) => Math.Round(value) + 0.5;
    }

    private sealed class Strip : FrameworkElement
    {
        private readonly SheetView _view;
        private readonly bool _horizontal;
        public Strip(SheetView view, bool horizontal) { _view = view; _horizontal = horizontal; ClipToBounds = true; }

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
                    bool on = column == v._selected.Column;
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
                    bool on = row == v._selected.Row;
                    var t = v.Text(row.ToString(CultureInfo.CurrentCulture), on ? App.B("AccentInk") : App.B("Ink2"), on);
                    dc.DrawText(t, new Point(w - 7 - t.Width, y + 4));
                }
                dc.DrawLine(line, new Point(Snap(w), 0), new Point(Snap(w), h));
            }
        }

        private static double Snap(double value) => Math.Round(value) + 0.5;
    }
}
