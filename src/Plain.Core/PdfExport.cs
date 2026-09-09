using System.Globalization;

namespace Plain.Core;

/// <summary>
/// Turning a file into a PDF. What Plain shows is what the PDF holds: the text of a document with its headings and
/// tables, the used part of each sheet, the words on each slide. It does not reproduce Word's pagination, because
/// Plain does not know it, and a PDF that claimed to would be a lie in a format people trust.
/// </summary>
public static class PdfExport
{
    public sealed record Result(byte[] Bytes, int Pages, string? Warning);

    public static Result Build(PlainFile file, string title) => Build(file, title, PageSetup.Default);

    public static Result Build(PlainFile file, string title, PageSetup page)
    {
        var pdf = page.NewPdf();
        string? warning = null;

        void Text(string text, double size, bool bold, double indent = 0, double after = 4)
        {
            if (!Pdf.CanWrite(text))
            {
                var missing = Pdf.Unwritable(text);
                warning ??= $"Some characters are not in the fonts a PDF reader already has ({Trim(missing)}); they are written as best they can be.";
            }
            pdf.Paragraph(text, size, bold, indent, after);
        }

        pdf.Paragraph(title, 15, bold: true, after: 2);
        pdf.Rule(56, 0, 0, 0, 0);   // no-op guard so a title with no body still yields a page
        pdf.Space(6);

        switch (file.Kind)
        {
            case FileKind.Document: Document(pdf, file.Document!, Text); break;
            case FileKind.Spreadsheet: Workbook(pdf, file.Workbook!, Text); break;
            case FileKind.Presentation: Deck(file.Deck!, Text, pdf); break;
        }

        var bytes = pdf.ToBytes();
        return new Result(bytes, Pages(bytes), warning);
    }

    private static string Trim(string s) => s.Length <= 12 ? s : s[..12] + "…";

    private static int Pages(byte[] pdf)
    {
        // The catalogue records it; counting "/Type /Page" would also catch /Pages.
        var text = System.Text.Encoding.ASCII.GetString(pdf);
        int at = text.IndexOf("/Count ", StringComparison.Ordinal);
        if (at < 0) return 1;
        var digits = new string(text.Skip(at + 7).TakeWhile(char.IsAsciiDigit).ToArray());
        return int.TryParse(digits, out var n) ? n : 1;
    }

    private static void Document(Pdf pdf, Document doc, Action<string, double, bool, double, double> text)
    {
        // Gather each table row before writing it, so its cells land side by side rather than one under the next.
        var blocks = doc.Blocks().ToList();
        int at = 0;
        while (at < blocks.Count)
        {
            var block = blocks[at];
            if (block.InTable)
            {
                int table = block.Table;
                // How wide is this table? The widest row decides, so every row lines up.
                int columns = blocks.Where(b => b.InTable && b.Table == table)
                                    .GroupBy(b => b.Row).Max(g => g.Count());
                double width = pdf.TextWidth / Math.Max(1, columns);

                while (at < blocks.Count && blocks[at].InTable && blocks[at].Table == table)
                {
                    int row = blocks[at].Row;
                    var cells = new List<(string, double, double, bool)>();
                    while (at < blocks.Count && blocks[at].InTable && blocks[at].Table == table && blocks[at].Row == row)
                    {
                        cells.Add((blocks[at].Text, 56 + cells.Count * width, width, false));
                        at++;
                    }
                    pdf.Row(cells, 9.5, bold: row == 0);
                }
                pdf.Space(6);
                continue;
            }
            at++;

            switch (block.Kind)
            {
                case BlockKind.Heading1: text(block.Text, 15, true, 0, 6); break;
                case BlockKind.Heading2: text(block.Text, 12.5, true, 0, 5); break;
                case BlockKind.Heading3: text(block.Text, 11, true, 0, 4); break;
                case BlockKind.ListItem: text("• " + block.Text, 10.5, false, 16, 2); break;
                default: text(block.Text, 10.5, false, 0, 5); break;
            }
        }
    }

    private static void Workbook(Pdf pdf, Workbook book, Action<string, double, bool, double, double> text)
    {
        bool first = true;
        foreach (var sheet in book.Sheets)
        {
            if (sheet.Hidden) continue;
            if (!first) pdf.NewPage();
            first = false;

            if (book.Sheets.Count > 1) text(sheet.Name, 12.5, true, 0, 4);

            var extent = sheet.Extent;
            int columns = Math.Min(extent.Column, 26);
            if (columns == 0) continue;

            // Give every column the same slice, which is what fits a sheet of unknown shape onto a page.
            double width = pdf.TextWidth / columns;
            for (int r = 1; r <= extent.Row; r++)
            {
                var cells = new List<(string, double, double, bool)>();
                for (int c = 1; c <= columns; c++)
                {
                    var cell = sheet.Read(new CellRef(c, r));
                    bool number = cell.Kind is CellKind.Number or CellKind.Formula
                                  && double.TryParse(cell.Raw, NumberStyles.Float, CultureInfo.InvariantCulture, out _);
                    cells.Add((cell.Display, 56 + (c - 1) * width, width, number));
                }
                pdf.Row(cells, 8.5, bold: r == 1);
            }

            if (extent.Column > columns)
                text($"Columns past {CellRef.ColumnName(columns)} are not on this page.", 8.5, false, 0, 2);
        }
    }

    private static void Deck(Deck deck, Action<string, double, bool, double, double> text, Pdf pdf)
    {
        bool first = true;
        foreach (var slide in deck.Slides)
        {
            if (!first) pdf.NewPage();
            first = false;
            text($"Slide {slide.Number}", 8.5, false, 0, 6);
            foreach (var frame in slide.Texts())
                foreach (var line in frame.Lines)
                {
                    if (line.Length == 0) continue;
                    bool title = frame.Placeholder is "title" or "ctrTitle";
                    text(line, title ? 16 : 11, title, title ? 0 : 14, title ? 8 : 3);
                }
        }
    }
}
