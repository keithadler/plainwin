namespace Plain.Core;

/// <summary>
/// Changing every occurrence of something. This is the single thing people ask for most: rebranding a policy for a
/// new client, correcting a code down a column, swapping a placeholder in a template. Doing it by hand with Find is
/// exactly how the previous client's name survives into a delivered document.
///
/// A spreadsheet replaces in what the cell holds, not in what it shows, so a number formatted as currency is matched
/// on its digits and a formula is matched on its formula. A cell whose whole content becomes a number stays a number.
/// </summary>
public static class Replace
{
    public sealed record Options(bool MatchCase = false, bool WholeCell = false, bool IncludeFormulas = true);

    public sealed record Result(int Cells, int Occurrences)
    {
        public bool Any => Occurrences > 0;
        public override string ToString() => $"{Occurrences} in {Cells}";
    }

    /// <summary>Replace inside one piece of text, counting how many times it happened.</summary>
    public static string InText(string text, string find, string with, Options options, out int count)
    {
        count = 0;
        if (find.Length == 0 || text.Length == 0) return text;

        var comparison = options.MatchCase ? StringComparison.Ordinal : StringComparison.CurrentCultureIgnoreCase;

        if (options.WholeCell)
        {
            if (!text.Equals(find, comparison)) return text;
            count = 1;
            return with;
        }

        var built = new System.Text.StringBuilder(text.Length);
        int at = 0;
        while (true)
        {
            int found = text.IndexOf(find, at, comparison);
            if (found < 0) break;
            built.Append(text, at, found - at).Append(with);
            at = found + find.Length;
            count++;
        }
        if (count == 0) return text;
        built.Append(text, at, text.Length - at);
        return built.ToString();
    }

    public static Result InWorkbook(Workbook book, string find, string with, Options options)
    {
        int cells = 0, occurrences = 0;
        foreach (var sheet in book.Sheets)
        {
            // Collect first: changing cells while walking them would be reading a list that is moving.
            var changes = new List<(CellRef Cell, string Value)>();
            foreach (var cell in sheet.Cells().ToList())
            {
                string source = cell.Formula ?? cell.Raw;
                if (cell.Formula is not null && !options.IncludeFormulas) continue;
                var replaced = InText(source, find, with, options, out int n);
                if (n == 0) continue;
                changes.Add((cell.Ref, replaced));
                cells++;
                occurrences += n;
            }
            foreach (var (cell, value) in changes) sheet.Set(cell, value);
        }
        return new Result(cells, occurrences);
    }

    public static Result InDocument(Document doc, string find, string with, Options options)
    {
        int blocks = 0, occurrences = 0;
        foreach (var block in doc.Blocks().ToList())
        {
            var replaced = InText(block.Text, find, with, options, out int n);
            if (n == 0) continue;
            doc.SetText(block.Index, replaced);
            blocks++;
            occurrences += n;
        }
        return new Result(blocks, occurrences);
    }

    public static Result InDeck(Deck deck, string find, string with, Options options)
    {
        int lines = 0, occurrences = 0;
        foreach (var slide in deck.Slides)
            foreach (var frame in slide.Texts().ToList())
                for (int line = 0; line < frame.Lines.Count; line++)
                {
                    var replaced = InText(frame.Lines[line], find, with, options, out int n);
                    if (n == 0) continue;
                    slide.SetLine(frame.ShapeIndex, line, replaced);
                    lines++;
                    occurrences += n;
                }
        return new Result(lines, occurrences);
    }

    /// <summary>Replace everywhere in whatever kind of file this is.</summary>
    public static Result InFile(PlainFile file, string find, string with, Options options) => file.Kind switch
    {
        FileKind.Spreadsheet => InWorkbook(file.Workbook!, find, with, options),
        FileKind.Document => InDocument(file.Document!, find, with, options),
        FileKind.Presentation => InDeck(file.Deck!, find, with, options),
        _ => new Result(0, 0),
    };
}
