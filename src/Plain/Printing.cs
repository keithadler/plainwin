using System.Windows;
using System.Windows.Documents;
using System.Windows.Media;
using Plain.Core;
using Table = System.Windows.Documents.Table;

namespace Plain;

/// <summary>
/// Putting a file on paper. Plain does not lay pages out the way Word does, so this prints what Plain shows rather
/// than pretending to reproduce somebody else's pagination: the text of a document, the used part of a sheet, the
/// words on each slide. It is honest about that, and it is enough for the worksheet, the invoice and the letter that
/// people said they could not do without.
/// </summary>
public static class Printing
{
    public static FlowDocument Build(PlainFile file, string title, double width, double height)
    {
        double margin = 48;
        var document = new FlowDocument
        {
            PageWidth = width,
            PageHeight = height,
            PagePadding = new Thickness(margin),
            ColumnWidth = width,               // one column, never newspaper columns
            FontFamily = new FontFamily("Segoe UI Variable Text, Segoe UI"),
            FontSize = 12,
            Foreground = Brushes.Black,
            Background = Brushes.White,
        };

        switch (file.Kind)
        {
            case FileKind.Document: Document(document, file.Document!); break;
            case FileKind.Spreadsheet: Workbook(document, file.Workbook!); break;
            case FileKind.Presentation: Deck(document, file.Deck!); break;
        }

        if (document.Blocks.Count == 0)
            document.Blocks.Add(new Paragraph(new Run($"{title} has nothing Plain can print.")));
        return document;
    }

    private static void Document(FlowDocument document, Core.Document source)
    {
        Table? table = null;
        int openTable = -1, openRow = -1;

        foreach (var block in source.Blocks())
        {
            if (block.InTable)
            {
                if (table is null || openTable != block.Table)
                {
                    table = new Table { CellSpacing = 0, Margin = new Thickness(0, 8, 0, 8) };
                    table.RowGroups.Add(new TableRowGroup());
                    document.Blocks.Add(table);
                    openTable = block.Table;
                    openRow = -1;
                }
                if (openRow != block.Row)
                {
                    table.RowGroups[0].Rows.Add(new TableRow());
                    openRow = block.Row;
                }
                while (table.Columns.Count <= block.Column) table.Columns.Add(new TableColumn());
                var cell = new TableCell(new Paragraph(new Run(block.Text)) { Margin = new Thickness(0) })
                {
                    BorderBrush = Brushes.DarkGray,
                    BorderThickness = new Thickness(0, 0, 1, 1),
                    Padding = new Thickness(5, 3, 5, 3),
                };
                table.RowGroups[0].Rows[^1].Cells.Add(cell);
                continue;
            }

            table = null;
            openTable = -1;

            var paragraph = new Paragraph(new Run(block.Text));
            switch (block.Kind)
            {
                case BlockKind.Heading1: paragraph.FontSize = 20; paragraph.FontWeight = FontWeights.Bold; paragraph.Margin = new Thickness(0, 14, 0, 6); break;
                case BlockKind.Heading2: paragraph.FontSize = 16; paragraph.FontWeight = FontWeights.Bold; paragraph.Margin = new Thickness(0, 12, 0, 5); break;
                case BlockKind.Heading3: paragraph.FontSize = 13; paragraph.FontWeight = FontWeights.Bold; paragraph.Margin = new Thickness(0, 10, 0, 4); break;
                case BlockKind.ListItem: paragraph.Margin = new Thickness(22, 2, 0, 2); break;
                default: paragraph.Margin = new Thickness(0, 3, 0, 3); break;
            }
            document.Blocks.Add(paragraph);
        }
    }

    private static void Workbook(FlowDocument document, Core.Workbook book)
    {
        bool several = book.Sheets.Count > 1;
        foreach (var sheet in book.Sheets)
        {
            if (sheet.Hidden) continue;
            var extent = sheet.Extent;
            if (several)
                document.Blocks.Add(new Paragraph(new Run(sheet.Name))
                {
                    FontSize = 15, FontWeight = FontWeights.Bold, Margin = new Thickness(0, 12, 0, 6),
                });

            var table = new Table { CellSpacing = 0, FontSize = 10.5 };
            table.RowGroups.Add(new TableRowGroup());
            for (int c = 1; c <= extent.Column; c++) table.Columns.Add(new TableColumn());
            document.Blocks.Add(table);

            for (int r = 1; r <= extent.Row; r++)
            {
                var row = new TableRow();
                for (int c = 1; c <= extent.Column; c++)
                {
                    var cell = sheet.Read(new CellRef(c, r));
                    bool number = cell.Kind is CellKind.Number or CellKind.Formula
                                  && double.TryParse(cell.Raw, System.Globalization.NumberStyles.Float,
                                                     System.Globalization.CultureInfo.InvariantCulture, out _);
                    row.Cells.Add(new TableCell(new Paragraph(new Run(cell.Display)) { Margin = new Thickness(0) })
                    {
                        BorderBrush = Brushes.LightGray,
                        BorderThickness = new Thickness(0, 0, 1, 1),
                        Padding = new Thickness(4, 2, 4, 2),
                        TextAlignment = number ? TextAlignment.Right : TextAlignment.Left,
                    });
                }
                table.RowGroups[0].Rows.Add(row);
            }
        }
    }

    private static void Deck(FlowDocument document, Core.Deck deck)
    {
        foreach (var slide in deck.Slides)
        {
            document.Blocks.Add(new Paragraph(new Run($"Slide {slide.Number}"))
            {
                FontSize = 10,
                Foreground = Brushes.Gray,
                Margin = new Thickness(0, 16, 0, 2),
            });
            foreach (var frame in slide.Texts())
                foreach (var line in frame.Lines)
                {
                    if (line.Length == 0) continue;
                    bool title = frame.Placeholder is "title" or "ctrTitle";
                    document.Blocks.Add(new Paragraph(new Run(line))
                    {
                        FontSize = title ? 17 : 12,
                        FontWeight = title ? FontWeights.Bold : FontWeights.Normal,
                        Margin = new Thickness(title ? 0 : 16, 2, 0, 2),
                    });
                }
        }
    }
}
