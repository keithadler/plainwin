namespace Plain.Core;

/// <summary>
/// The shape a block of cells takes on the clipboard: rows of tab separated values. Every spreadsheet reads and
/// writes this, so it is how cells move between Plain, Excel, a text editor and a browser.
/// </summary>
public static class Tabular
{
    public static string Write(IReadOnlyList<IReadOnlyList<string>> rows)
    {
        var text = new System.Text.StringBuilder();
        for (int r = 0; r < rows.Count; r++)
        {
            if (r > 0) text.Append("\r\n");
            for (int c = 0; c < rows[r].Count; c++)
            {
                if (c > 0) text.Append('\t');
                text.Append(rows[r][c]);
            }
        }
        return text.ToString();
    }

    /// <summary>
    /// Read a block back. Line endings come in every combination depending on where the text was copied from, one
    /// trailing newline is normal and must not become an empty row, and rows are allowed to be ragged.
    /// </summary>
    public static IReadOnlyList<IReadOnlyList<string>> Read(string text)
    {
        if (string.IsNullOrEmpty(text)) return Array.Empty<IReadOnlyList<string>>();
        var normalised = text.Replace("\r\n", "\n").Replace('\r', '\n');
        // One trailing newline ends the last row rather than starting an empty one.
        if (normalised.EndsWith('\n')) normalised = normalised[..^1];
        return normalised.Split('\n').Select(line => (IReadOnlyList<string>)line.Split('\t')).ToList();
    }

    /// <summary>The widest row, which is how far a paste reaches across.</summary>
    public static int Width(IReadOnlyList<IReadOnlyList<string>> rows) =>
        rows.Count == 0 ? 0 : rows.Max(r => r.Count);
}
