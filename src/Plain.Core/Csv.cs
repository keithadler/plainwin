using System.Globalization;
using System.Text;

namespace Plain.Core;

/// <summary>
/// Comma separated values, in and out. Every pipeline in an office ends or begins in one, and a spreadsheet tool
/// that cannot read or write one cannot be in the pipeline at all.
///
/// Reading follows the usual rules: quotes protect commas and newlines, a doubled quote is a quote. Writing quotes
/// only what needs it. A cell that holds a formula is written as its worked-out value, because that is what anything
/// reading a CSV expects, and the formula is said to be there rather than silently dropped.
/// </summary>
public static class Csv
{
    public static string Write(Sheet sheet, char separator = ',')
    {
        var extent = sheet.Extent;
        var built = new StringBuilder();
        for (int r = 1; r <= extent.Row; r++)
        {
            for (int c = 1; c <= extent.Column; c++)
            {
                if (c > 1) built.Append(separator);
                built.Append(Quote(sheet.Read(new CellRef(c, r)).Display, separator));
            }
            built.Append("\r\n");
        }
        return built.ToString();
    }

    public static string Quote(string field, char separator = ',')
    {
        bool needs = field.Contains(separator) || field.Contains('"') || field.Contains('\n') || field.Contains('\r')
                     || field.StartsWith(' ') || field.EndsWith(' ');
        if (!needs) return field;
        return '"' + field.Replace("\"", "\"\"") + '"';
    }

    /// <summary>Read a whole CSV into rows of fields.</summary>
    public static IReadOnlyList<IReadOnlyList<string>> Read(string text, char separator = ',')
    {
        var rows = new List<IReadOnlyList<string>>();
        var row = new List<string>();
        var field = new StringBuilder();
        bool quoted = false, any = false;

        for (int i = 0; i < text.Length; i++)
        {
            char c = text[i];
            if (quoted)
            {
                if (c == '"')
                {
                    if (i + 1 < text.Length && text[i + 1] == '"') { field.Append('"'); i++; }
                    else quoted = false;
                }
                else field.Append(c);
                continue;
            }

            if (c == '"' && field.Length == 0) { quoted = true; any = true; continue; }
            if (c == separator) { row.Add(field.ToString()); field.Clear(); any = true; continue; }
            if (c is '\r') continue;
            if (c is '\n')
            {
                row.Add(field.ToString());
                rows.Add(row);
                row = new List<string>();
                field.Clear();
                any = false;
                continue;
            }
            field.Append(c);
            any = true;
        }
        if (any || field.Length > 0 || row.Count > 0) { row.Add(field.ToString()); rows.Add(row); }
        return rows;
    }

    /// <summary>
    /// Put a CSV into a sheet, starting at the given corner. Values that look like numbers become numbers, so a
    /// column of prices arrives as prices; anything beginning with = is left as text, because a CSV should never be
    /// able to put a formula into somebody's workbook.
    /// </summary>
    public static int Into(Sheet sheet, string csv, CellRef at, char separator = ',')
    {
        var rows = Read(csv, separator);
        int written = 0;
        for (int r = 0; r < rows.Count; r++)
            for (int c = 0; c < rows[r].Count; c++)
            {
                var value = rows[r][c];
                if (value.StartsWith('=')) value = "'" + value;   // never let a file inject a formula
                sheet.Set(new CellRef(at.Column + c, at.Row + r), value);
                written++;
            }
        return written;
    }
}
