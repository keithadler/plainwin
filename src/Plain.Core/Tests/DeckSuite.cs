namespace Plain.Core.Tests;

public static class DeckSuite
{
    public static Suite Run()
    {
        var s = new Suite("deck");
        if (Fixtures.Missing(s, "deck.pptx")) return s;
        var path = Fixtures.Path_("deck.pptx")!;

        var deck = new Deck(OpcPackage.Open(path));
        s.Equal("two slides", 2, deck.Slides.Count);
        s.Equal("slides are numbered from one", 1, deck.Slides[0].Number);
        s.Equal("the first slide's title", "Kickoff", deck.Slides[0].Title());
        s.Equal("the second slide's title", "Where we are today", deck.Slides[1].Title());
        s.Check("the second slide has body lines",
                deck.Slides[1].Texts().Any(t => t.Lines.Any(l => l.Contains("Month-end close"))));

        var work = Fixtures.Copy("deck.pptx");
        try
        {
            var d = new Deck(OpcPackage.Open(work));
            var slide = d.Slides[1];
            var body = slide.Texts().First(t => t.Lines.Any(l => l.Contains("Month-end close")));
            int line = body.Lines.ToList().FindIndex(l => l.Contains("Month-end close"));
            string neighbour = body.Lines[0];
            slide.SetLine(body.ShapeIndex, line, "Month-end close takes four working days");
            d.Save(work);

            var again = new Deck(OpcPackage.Open(work));
            var after = again.Slides[1].Texts().First(t => t.ShapeIndex == body.ShapeIndex);
            s.Equal("the edited line reads back", "Month-end close takes four working days", after.Lines[line]);
            s.Equal("a neighbouring line is untouched", neighbour, after.Lines[0]);
            s.Equal("the title is untouched", "Where we are today", again.Slides[1].Title());
            s.Equal("the other slide is untouched", "Kickoff", again.Slides[0].Title());
            s.Equal("the line count is unchanged", body.Lines.Count, after.Lines.Count);
        }
        finally { try { File.Delete(work); } catch { } }

        return s;
    }
}
