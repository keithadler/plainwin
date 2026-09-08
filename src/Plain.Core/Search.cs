namespace Plain.Core;

/// <summary>
/// Finding text, in the order someone reads the file: across the sheets of a workbook, down the blocks of a document,
/// through the slides of a deck. A spreadsheet search looks at what the cell shows and at the formula behind it, so
/// searching for a function name finds the cells that use it.
/// </summary>
public static class Search
{
    public static bool Matches(string haystack, string term) =>
        haystack.Contains(term, StringComparison.CurrentCultureIgnoreCase);

    public sealed record CellHit(int Sheet, CellRef Cell, string Where, string Text);
    public sealed record BlockHit(int Block, string Where, string Text);
    public sealed record SlideHit(int Slide, int Shape, int Line, string Where, string Text);

    public static IReadOnlyList<CellHit> InWorkbook(Workbook book, string term)
    {
        var hits = new List<CellHit>();
        if (term.Length == 0) return hits;
        for (int i = 0; i < book.Sheets.Count; i++)
        {
            var sheet = book.Sheets[i];
            foreach (var cell in sheet.Cells())
            {
                if (!Matches(cell.Display, term) && !Matches(cell.Formula ?? "", term)) continue;
                hits.Add(new CellHit(i, cell.Ref, $"{sheet.Name}!{cell.Ref}", cell.Formula ?? cell.Display));
            }
        }
        return hits;
    }

    public static IReadOnlyList<BlockHit> InDocument(Document doc, string term)
    {
        var hits = new List<BlockHit>();
        if (term.Length == 0) return hits;
        foreach (var block in doc.Blocks())
            if (Matches(block.Text, term))
                hits.Add(new BlockHit(block.Index, $"block {block.Index}", block.Text));
        return hits;
    }

    public static IReadOnlyList<SlideHit> InDeck(Deck deck, string term)
    {
        var hits = new List<SlideHit>();
        if (term.Length == 0) return hits;
        foreach (var slide in deck.Slides)
            foreach (var frame in slide.Texts())
                for (int line = 0; line < frame.Lines.Count; line++)
                    if (Matches(frame.Lines[line], term))
                        hits.Add(new SlideHit(slide.Number, frame.ShapeIndex, line,
                                              $"slide {slide.Number}", frame.Lines[line]));
        return hits;
    }
}
