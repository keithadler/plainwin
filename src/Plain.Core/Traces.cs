namespace Plain.Core;

/// <summary>
/// What a formula reads, and what reads a cell.
///
/// This is the question behind most spreadsheet mistakes. A total looks wrong and you cannot see which cells it is
/// adding up; you change a number and cannot see what else moved because of it. Both answers are in the file
/// already, in the formulas, and Plain can read them, so it may as well say.
///
/// It answers about the file as it is written, not as Excel would compute it: a reference to a whole column is
/// reported as a whole column, and a name Plain does not understand is reported as not understood rather than
/// guessed at.
/// </summary>
public static class Traces
{
    /// <summary>One cell that takes part, and how.</summary>
    public sealed record Touch(string Where, string What);

    /// <summary>What the cell's own formula reads. Empty when the cell holds a value rather than a formula.</summary>
    public static IReadOnlyList<Touch> Reads(Workbook book, Sheet sheet, CellRef cell)
    {
        var found = new List<Touch>();
        var formula = sheet.Read(cell).Formula;
        if (formula is null) return found;

        // The formula as Cell reports it carries its leading sign; the scanner wants the text after it.
        var text = formula.StartsWith('=') ? formula[1..] : formula;

        foreach (var token in Refs.Scan(text))
        {
            var range = token.Range;
            var where = range.Sheet is { Length: > 0 } named ? named : sheet.Name;
            var what = range.RowMin == range.RowMax && range.ColumnMin == range.ColumnMax
                ? new CellRef(range.ColumnMin, range.RowMin).ToString()
                : $"{new CellRef(range.ColumnMin, range.RowMin)}:{new CellRef(range.ColumnMax, range.RowMax)}";

            var target = book.Sheets.FirstOrDefault(x => x.Name.Equals(where, StringComparison.OrdinalIgnoreCase));
            string says = target is null
                ? "on a sheet that is not in this workbook"
                : range.RowMin == range.RowMax && range.ColumnMin == range.ColumnMax
                    ? Describe(target.Read(new CellRef(range.ColumnMin, range.RowMin)))
                    : $"{Filled(target, range)} cells with something in them";

            found.Add(new Touch($"{where}!{what}", says));
        }
        return found;
    }

    /// <summary>Every formula anywhere in the workbook that reads this cell.</summary>
    public static IReadOnlyList<Touch> ReadBy(Workbook book, Sheet sheet, CellRef cell)
    {
        var found = new List<Touch>();
        foreach (var other in book.Sheets)
            foreach (var (at, formula) in other.Formulas())
            {
                bool touches = false;
                foreach (var token in Refs.Scan(formula))
                {
                    var named = token.Range.Sheet;
                    bool same = named is null or "" ? ReferenceEquals(other, sheet)
                                                    : named.Equals(sheet.Name, StringComparison.OrdinalIgnoreCase);
                    if (!same) continue;
                    var range = token.Range;
                    if (cell.Column < range.ColumnMin || cell.Column > range.ColumnMax) continue;
                    if (cell.Row < range.RowMin || cell.Row > range.RowMax) continue;
                    touches = true;
                    break;
                }
                if (touches) found.Add(new Touch($"{other.Name}!{at}", "=" + Shorten(formula)));
            }
        return found;
    }

    private static string Describe(Cell cell) => cell.Kind switch
    {
        CellKind.Empty => "empty",
        CellKind.Formula => "another formula: =" + Shorten(cell.Formula ?? ""),
        _ => cell.Display.Length > 40 ? cell.Display[..40] + "…" : cell.Display,
    };

    private static int Filled(Sheet sheet, RefRange range)
    {
        int n = 0;
        // A whole-column reference would be a million reads; count what the sheet actually reaches.
        var extent = sheet.Extent;
        int lastRow = Math.Min(range.RowMax, Math.Max(extent.Row, 1));
        int lastColumn = Math.Min(range.ColumnMax, Math.Max(extent.Column, 1));
        for (int r = range.RowMin; r <= lastRow; r++)
            for (int c = range.ColumnMin; c <= lastColumn; c++)
                if (sheet.Read(new CellRef(c, r)).Kind != CellKind.Empty) n++;
        return n;
    }

    private static string Shorten(string formula)
    {
        var one = formula.StartsWith('=') ? formula[1..] : formula;
        return one.Length <= 60 ? one : one[..60] + "…";
    }
}
