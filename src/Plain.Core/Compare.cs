namespace Plain.Core;

/// <summary>
/// What changed between two versions of a file.
///
/// This is the question you have when a document comes back from someone: what did they actually touch. Plain can
/// answer it in a way most tools cannot, because it already holds a file as the parts it is made of. Comparing the
/// parts says which pieces of the file differ at all, and then the text inside the pieces Plain understands says
/// what differs in words.
///
/// It compares what is there; it does not try to work out how one became the other. A line that moved reads as one
/// gone and one arrived, and saying so plainly is better than guessing at a rename and being wrong.
/// </summary>
public static class Compare
{
    /// <summary>One part of the file, and how the two versions of it differ.</summary>
    public sealed record PartChange(string Name, string How);

    /// <summary>One line of text that is in one version and not the other.</summary>
    public sealed record TextChange(string Where, string Text, bool Added);

    public sealed record Report(
        IReadOnlyList<PartChange> Parts,
        IReadOnlyList<TextChange> Text,
        int Same,
        string? Note);

    /// <summary>Compare two files. They do not have to be the same kind, but a useful answer needs them to be.</summary>
    public static Report Between(PlainFile before, PlainFile after)
    {
        var parts = new List<PartChange>();
        var text = new List<TextChange>();
        int same = 0;

        var beforeNames = before.Package.Parts.Select(p => p.Name).ToHashSet(StringComparer.Ordinal);
        var afterNames = after.Package.Parts.Select(p => p.Name).ToHashSet(StringComparer.Ordinal);

        foreach (var name in beforeNames.Union(afterNames).OrderBy(x => x, StringComparer.Ordinal))
        {
            bool inBefore = beforeNames.Contains(name), inAfter = afterNames.Contains(name);
            if (!inAfter) { parts.Add(new PartChange(name, "gone")); continue; }
            if (!inBefore) { parts.Add(new PartChange(name, "new")); continue; }

            // The bytes are what Plain holds, so this is exact rather than a guess about meaning.
            if (before.Package.Read(name).AsSpan().SequenceEqual(after.Package.Read(name))) { same++; continue; }
            parts.Add(new PartChange(name, "changed"));
        }

        string? note = null;
        if (before.Kind != after.Kind)
            note = "These are different kinds of file, so only the parts can be compared.";
        else
        {
            var was = Lines(before);
            var now = Lines(after);
            // Counting each line means a line that appears twice and then once reads as one gone, which is right.
            var wasCount = Count(was);
            var nowCount = Count(now);

            foreach (var (line, n) in wasCount)
            {
                int still = nowCount.TryGetValue(line, out var m) ? m : 0;
                for (int i = 0; i < n - still && text.Count < 500; i++)
                    text.Add(new TextChange(line.Where, line.Text, Added: false));
            }
            foreach (var (line, n) in nowCount)
            {
                int had = wasCount.TryGetValue(line, out var m) ? m : 0;
                for (int i = 0; i < n - had && text.Count < 500; i++)
                    text.Add(new TextChange(line.Where, line.Text, Added: true));
            }

            if (text.Count >= 500) note = "There were too many changes to list them all; these are the first 500.";
        }

        return new Report(parts, text, same, note);
    }

    private sealed record Line(string Where, string Text);

    private static Dictionary<Line, int> Count(IEnumerable<Line> lines)
    {
        var map = new Dictionary<Line, int>();
        foreach (var line in lines) map[line] = map.TryGetValue(line, out var n) ? n + 1 : 1;
        return map;
    }

    /// <summary>
    /// Every line of text Plain can read from a file, each with where it came from. Where matters: the same words
    /// in a different cell are a change, and saying only that the words are still somewhere would hide it.
    /// </summary>
    private static IEnumerable<Line> Lines(PlainFile file)
    {
        if (file.Workbook is { } book)
            foreach (var sheet in book.Sheets)
            {
                var extent = sheet.Extent;
                for (int r = 1; r <= extent.Row; r++)
                    for (int c = 1; c <= extent.Column; c++)
                    {
                        var cell = sheet.Read(new CellRef(c, r));
                        if (cell.Kind == CellKind.Empty) continue;
                        yield return new Line($"{sheet.Name}!{new CellRef(c, r)}", cell.Formula ?? cell.Display);
                    }
            }
        else if (file.Document is { } doc)
            foreach (var block in doc.Blocks())
            {
                if (block.Text.Trim().Length == 0) continue;
                yield return new Line("the text", block.Text);
            }
        else if (file.Deck is { } deck)
            foreach (var slide in deck.Slides)
                foreach (var body in slide.Texts())
                    foreach (var line in body.Lines)
                    {
                        if (line.Trim().Length == 0) continue;
                        yield return new Line($"slide {slide.Number}", line);
                    }
    }
}
