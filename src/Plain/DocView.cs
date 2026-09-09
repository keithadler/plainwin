using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Data;
using System.Windows.Input;
using System.Windows.Media;
using Plain.Core;
// Plain.Core has its own Grid, for rows and columns of a sheet; in this file Grid means the WPF panel.
using Grid = System.Windows.Controls.Grid;
using Search = Plain.Core.Search;

namespace Plain;

/// <summary>
/// A Word document as one scrolling column of editable text. Plain does not lay pages out the way Word does, because
/// matching Word's pagination needs Word's own fonts and line breaking; the status bar says so rather than pretending.
/// What it does do is let you fix the words without touching anything else in the file, and show a table as a table.
/// </summary>
public sealed class DocView : Grid, IFindable
{
    private readonly Document _doc;
    private readonly ObservableCollection<object> _rows = new();

    public event Action<Action>? Edited;

    /// <summary>The block the caret is in, so a format button knows what to act on.</summary>
    public int? FocusedBlock { get; private set; }

    /// <summary>True once an edit flattened mixed formatting inside one block, so the app can say so honestly.</summary>
    public bool FlattenedSomething { get; private set; }

    public DocView(Document doc)
    {
        _doc = doc;
        BuildRows();

        _list = new ItemsControl
        {
            ItemsSource = _rows,
            Margin = new Thickness(44, 26, 44, 60),
            MaxWidth = 860,
            HorizontalAlignment = HorizontalAlignment.Left,
            ItemTemplateSelector = new RowTemplates(),
        };
        // One shared size scope keeps a table's columns lined up from one row to the next.
        _list.SetValue(Grid.IsSharedSizeScopeProperty, true);
        _list.SetValue(VirtualizingStackPanel.IsVirtualizingProperty, true);
        _list.SetValue(VirtualizingStackPanel.VirtualizationModeProperty, VirtualizationMode.Recycling);

        _scroller = new ScrollViewer
        {
            Content = _list,
            VerticalScrollBarVisibility = ScrollBarVisibility.Auto,
            HorizontalScrollBarVisibility = ScrollBarVisibility.Disabled,
            Background = App.B("Surface"),
            CanContentScroll = false,
            Focusable = false,
        };
        Children.Add(_scroller);
    }

    private readonly ItemsControl _list;
    private readonly ScrollViewer _scroller;
    private IReadOnlyList<Search.BlockHit> _hits = Array.Empty<Search.BlockHit>();

    public int FindAll(string term)
    {
        _hits = Search.InDocument(_doc, term);
        return _hits.Count;
    }

    public string Reveal(int index)
    {
        if (index < 0 || index >= _hits.Count) return "";
        var line = Find(_hits[index].Block);
        if (line is null) return "";

        // Work out which row holds the block, then scroll roughly there so the virtualized panel builds it,
        // and only then ask the realized container to bring itself fully into view.
        object? row = _rows.FirstOrDefault(r => r == line || (r is TableRow t && t.Cells.Contains(line)));
        if (row is null) return "";
        int position = _rows.IndexOf(row);

        _scroller.ScrollToVerticalOffset(_scroller.ExtentHeight * position / Math.Max(1, _rows.Count));
        _scroller.UpdateLayout();
        if (_list.ItemContainerGenerator.ContainerFromItem(row) is FrameworkElement container)
        {
            container.BringIntoView();
            var box = FirstTextBox(container);
            box?.Focus();
            box?.SelectAll();
        }
        return $"block {_hits[index].Block}";
    }

    private static TextBox? FirstTextBox(DependencyObject root)
    {
        if (root is TextBox box) return box;
        int count = VisualTreeHelper.GetChildrenCount(root);
        for (int i = 0; i < count; i++)
            if (FirstTextBox(VisualTreeHelper.GetChild(root, i)) is { } found) return found;
        return null;
    }

    /// <summary>Group the document's blocks so that the cells of one table row become one item.</summary>
    private void BuildRows()
    {
        _rows.Clear();
        TableRow? open = null;
        foreach (var block in _doc.Blocks())
        {
            var line = new BlockLine(this, block);
            if (!block.InTable) { open = null; _rows.Add(line); continue; }

            if (open is null || open.Table != block.Table || open.Row != block.Row)
            {
                open = new TableRow(block.Table, block.Row);
                _rows.Add(open);
            }
            open.Cells.Add(line);
        }
    }

    internal void Apply(int index, string text, string was)
    {
        if (!_doc.SetText(index, text)) FlattenedSomething = true;
        Edited?.Invoke(() =>
        {
            _doc.SetText(index, was);
            Find(index)?.Restore(was);
        });
    }

    private BlockLine? Find(int index)
    {
        foreach (var row in _rows)
        {
            if (row is BlockLine line && line.Index == index) return line;
            if (row is TableRow table)
                foreach (var cell in table.Cells) if (cell.Index == index) return cell;
        }
        return null;
    }

    /// <summary>One row of a table, so its cells sit side by side instead of stacking.</summary>
    public sealed class TableRow
    {
        public int Table { get; }
        public int Row { get; }
        public List<BlockLine> Cells { get; } = new();
        public TableRow(int table, int row) { Table = table; Row = row; }
    }

    /// <summary>One block, bound to a text box. Setting Text writes straight through to the document.</summary>
    public sealed class BlockLine : INotifyPropertyChanged
    {
        private readonly DocView _view;
        private string _text;

        public int Index { get; }
        public BlockKind Kind { get; }

        public BlockLine(DocView view, Block block)
        {
            _view = view; Index = block.Index; Kind = block.Kind; _text = block.Text;
        }

        public string Text
        {
            get => _text;
            set
            {
                if (value == _text) return;
                string was = _text;
                _text = value;
                _view.Apply(Index, value, was);
                PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(Text)));
            }
        }

        /// <summary>Put the old text back without writing it through again, which is what undo needs.</summary>
        internal void Restore(string text)
        {
            _text = text;
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(Text)));
        }

        public double FontSize => Kind switch
        {
            BlockKind.Heading1 => 21,
            BlockKind.Heading2 => 16,
            BlockKind.Heading3 => 14,
            _ => 13.5,
        };

        public FontWeight Weight => Kind is BlockKind.Heading1 or BlockKind.Heading2 or BlockKind.Heading3
            ? FontWeights.SemiBold : FontWeights.Normal;

        public Thickness Indent => Kind switch
        {
            BlockKind.Heading1 => new Thickness(0, 14, 0, 4),
            BlockKind.Heading2 => new Thickness(0, 16, 0, 3),
            BlockKind.Heading3 => new Thickness(0, 12, 0, 2),
            BlockKind.ListItem => new Thickness(22, 1, 0, 1),
            _ => new Thickness(0, 2, 0, 2),
        };

        public event PropertyChangedEventHandler? PropertyChanged;
    }

    /// <summary>A paragraph gets a text box; a table row gets a grid of them.</summary>
    private sealed class RowTemplates : DataTemplateSelector
    {
        private readonly DataTemplate _paragraph = (DataTemplate)System.Windows.Markup.XamlReader.Parse(
            "<DataTemplate xmlns='http://schemas.microsoft.com/winfx/2006/xaml/presentation'>" +
            "<TextBox Text='{Binding Text, Mode=TwoWay, UpdateSourceTrigger=LostFocus}' " +
            "         FontSize='{Binding FontSize}' FontWeight='{Binding Weight}' Margin='{Binding Indent}' " +
            "         BorderThickness='0' Background='Transparent' Padding='2,1' " +
            "         Foreground='{DynamicResource Ink}' TextWrapping='Wrap' AcceptsReturn='False' " +
            "         SpellCheck.IsEnabled='True'/></DataTemplate>");

        public override DataTemplate? SelectTemplate(object item, DependencyObject container) =>
            item is TableRow ? BuildTableRow() : _paragraph;

        private static DataTemplate BuildTableRow()
        {
            // Built in code because the number of columns is only known per row.
            var template = new DataTemplate(typeof(TableRow));
            var factory = new FrameworkElementFactory(typeof(TableRowPanel));
            template.VisualTree = factory;
            template.Seal();
            return template;
        }
    }

    /// <summary>Lays one table row out as a grid whose columns line up with the rest of its table.</summary>
    private sealed class TableRowPanel : Grid
    {
        public TableRowPanel()
        {
            DataContextChanged += (_, _) => Rebuild();
            Margin = new Thickness(0, 1, 0, 1);
        }

        private void Rebuild()
        {
            Children.Clear();
            ColumnDefinitions.Clear();
            if (DataContext is not TableRow row) return;

            for (int i = 0; i < row.Cells.Count; i++)
            {
                ColumnDefinitions.Add(new ColumnDefinition
                {
                    Width = GridLength.Auto,
                    MinWidth = 90,
                    SharedSizeGroup = $"t{row.Table}c{i}",
                });

                var box = new TextBox
                {
                    FontSize = 12.5,
                    BorderBrush = App.B("Line"),
                    BorderThickness = new Thickness(0, 0, 1, 1),
                    Background = Brushes.Transparent,
                    Foreground = App.B("Ink"),
                    Padding = new Thickness(7, 3, 7, 3),
                    TextWrapping = TextWrapping.Wrap,
                    AcceptsReturn = false,
                };
                box.SetBinding(TextBox.TextProperty, new Binding("Text")
                {
                    Source = row.Cells[i],
                    Mode = BindingMode.TwoWay,
                    UpdateSourceTrigger = UpdateSourceTrigger.LostFocus,
                });
                SetColumn(box, i);
                Children.Add(box);
            }
        }
    }
}
