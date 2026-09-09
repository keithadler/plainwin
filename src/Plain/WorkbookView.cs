using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using Plain.Core;
// Plain.Core has its own Grid, for rows and columns of a sheet; in this file Grid means the WPF panel.
using Grid = System.Windows.Controls.Grid;

namespace Plain;

/// <summary>
/// A whole workbook: the grid for the sheet you are on, and a strip of the other sheets along the bottom where Excel
/// puts them. Each sheet keeps its own grid once you have visited it, so switching back is instant and your selection
/// is where you left it.
/// </summary>
public sealed class WorkbookView : Grid, IFindable
{
    private readonly Workbook _book;
    private readonly Dictionary<Sheet, SheetView> _grids = new();
    private readonly ContentControl _host = new();
    private readonly StackPanel _strip = new() { Orientation = Orientation.Horizontal };
    private SheetView _current = null!;

    public event Action<CellRef, Cell>? SelectionChanged;
    public event Action<Edit>? Edited;
    public event Action<GridEdit, int>? GridChangeRequested;
    public event Action<int, int, int, int, int, bool>? SortRequested;
    public event Action<int, int>? FilterRequested;

    /// <summary>Asked to do something to a sheet rather than in one; the window owns the questions and the undo.</summary>
    public event Action<string, int>? SheetChangeRequested;

    public bool Filtering => _current.Filtering;
    public string FilterSaid => _current.FilterSaid;
    public void Filter(int column, string text, int firstRow) => _current.Filter(column, text, firstRow);
    public void ClearFilter() => _current.ClearFilter();

    public Sheet CurrentSheet => _current.Sheet;

    /// <summary>What the selection adds up to, for the status bar.</summary>
    public string Summary() => _current.Summary();

    /// <summary>Go to a cell by name.</summary>
    public bool GoTo(string reference) => _current.GoTo(reference);

    /// <summary>Draw the sheet again after something changed underneath it, and forget what was cached.</summary>
    public void Redraw() => _current.Reload();
    public SheetView CurrentGrid => _current;

    public WorkbookView(Workbook book)
    {
        _book = book;
        RowDefinitions.Add(new RowDefinition { Height = new GridLength(1, GridUnitType.Star) });
        RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });

        SetRow(_host, 0);
        Children.Add(_host);

        // A workbook with one sheet needs no strip; showing an empty one would be furniture for its own sake.
        if (book.Sheets.Count > 1)
        {
            var scroller = new ScrollViewer
            {
                Content = _strip,
                HorizontalScrollBarVisibility = ScrollBarVisibility.Auto,
                VerticalScrollBarVisibility = ScrollBarVisibility.Disabled,
                Background = App.B("Surface2"),
            };
            var bar = new Border
            {
                Child = scroller,
                BorderBrush = App.B("Line"),
                BorderThickness = new Thickness(0, 1, 0, 0),
                Padding = new Thickness(6, 3, 6, 3),
            };
            SetRow(bar, 1);
            Children.Add(bar);
        }

        // Start on the first sheet anyone can see; a workbook where every sheet is hidden still has to open.
        Show(book.Sheets.FirstOrDefault(s => !s.Hidden) ?? book.Sheets[0]);
    }

    private void Show(Sheet sheet)
    {
        if (!_grids.TryGetValue(sheet, out var grid))
        {
            grid = new SheetView(sheet);
            grid.SelectionChanged += (reference, cell) => SelectionChanged?.Invoke(reference, cell);
            grid.Edited += edit => Edited?.Invoke(edit);
            grid.GridChangeRequested += (edit, at) => GridChangeRequested?.Invoke(edit, at);
            grid.SortRequested += (t, b, l, r, key, up) => SortRequested?.Invoke(t, b, l, r, key, up);
            grid.FilterRequested += (column, row) => FilterRequested?.Invoke(column, row);
            _grids[sheet] = grid;
        }
        _current = grid;
        _host.Content = grid;
        BuildStrip();
        Dispatcher.BeginInvoke(new Action(() => { grid.Focus(); grid.Redraw(); }),
            System.Windows.Threading.DispatcherPriority.Loaded);
        SelectionChanged?.Invoke(grid.Selected, grid.Sheet.Read(grid.Selected));
    }

    private void BuildStrip()
    {
        _strip.Children.Clear();
        foreach (var sheet in _book.Sheets)
        {
            bool on = sheet == _current.Sheet;
            var tab = new Button
            {
                Content = new TextBlock
                {
                    Text = sheet.Name + (sheet.Hidden ? "  (hidden)" : ""),
                    FontSize = 12,
                    FontWeight = on ? FontWeights.SemiBold : FontWeights.Normal,
                    Foreground = on ? App.B("AccentInk") : App.B("Ink2"),
                },
                Padding = new Thickness(11, 4, 11, 4),
                Margin = new Thickness(0, 0, 3, 0),
                Background = on ? App.B("Surface") : Brushes.Transparent,
                BorderBrush = on ? App.B("Accent") : Brushes.Transparent,
                BorderThickness = new Thickness(0, 2, 0, 0),
                Cursor = Cursors.Hand,
                Tag = sheet,
            };
            tab.Click += (s, _) => Show((Sheet)((Button)s).Tag);

            // Right-click a tab for the things you do to a sheet rather than in one.
            var menu = new ContextMenu();
            void Item(string text, Action<int> what)
            {
                var entry = new MenuItem { Header = text };
                entry.Click += (_, _) => what(_book.Sheets.ToList().IndexOf((Sheet)tab.Tag));
                menu.Items.Add(entry);
            }
            Item("Rename this sheet…", i => SheetChangeRequested?.Invoke("rename", i));
            Item("Add a sheet after this one", i => SheetChangeRequested?.Invoke("add", i));
            Item("Move it earlier", i => SheetChangeRequested?.Invoke("earlier", i));
            Item("Move it later", i => SheetChangeRequested?.Invoke("later", i));
            menu.Items.Add(new Separator());
            Item("Take this sheet out…", i => SheetChangeRequested?.Invoke("remove", i));
            tab.ContextMenu = menu;
            _strip.Children.Add(tab);
        }
    }

    /// <summary>Put a typed value into the selected cell of the sheet on screen.</summary>
    public void Apply(string typed) => _current.Apply(typed);

    private IReadOnlyList<Search.CellHit> _hits = Array.Empty<Search.CellHit>();

    public int FindAll(string term)
    {
        _hits = Search.InWorkbook(_book, term);
        return _hits.Count;
    }

    public string Reveal(int index)
    {
        if (index < 0 || index >= _hits.Count) return "";
        var hit = _hits[index];
        var sheet = _book.Sheets[hit.Sheet];
        if (sheet != _current.Sheet) Show(sheet);
        _grids[sheet].Select(hit.Cell);
        return hit.Where;
    }
}
