namespace Plain.Core.Tests;

/// <summary>
/// Moving, removing and adding slides. This writes new parts into a package, which is the most dangerous thing
/// Plain does, so every check ends by opening the file again: a deck that cannot be reopened is a deck that has
/// been broken, whatever the operation said.
/// </summary>
public static class SlidesSuite
{
    private static string Order(PlainFile file) =>
        string.Join(",", file.Deck!.Slides.Select(s => s.Title()));

    private static PlainFile Reopen(PlainFile file)
    {
        file.Flush();
        return PlainFile.Read(file.Package.ToBytes(), file.Path);
    }

    public static Suite Run()
    {
        var s = new Suite("slides");

        var corpus = Environment.GetEnvironmentVariable("PLAIN_CORPUS");
        var demo = corpus is not null ? System.IO.Path.Combine(corpus, "woodland.pptx") : null;
        if (demo is null || !File.Exists(demo))
        {
            if (Fixtures.Missing(s, "deck.pptx")) return s;
            demo = Fixtures.Path_("deck.pptx");
        }

        string Work()
        {
            var path = System.IO.Path.Combine(System.IO.Path.GetTempPath(),
                "plain-slides-" + Guid.NewGuid().ToString("N") + ".pptx");
            File.Copy(demo!, path, overwrite: true);
            return path;
        }

        // ---- moving ----
        {
            var file = PlainFile.Open(Work());
            int count = file.Deck!.Slides.Count;
            var was = Order(file);

            var moved = Slides.Move(file.Package, 1, count);
            s.Check("moving a slide says it did", moved is Slides.Done);
            var after = Reopen(file);
            s.Check("the deck still opens after a move", after.Deck is not null);
            s.Check("there are still the same number of slides", after.Deck!.Slides.Count == count);
            s.Check("and the order really changed", Order(after) != was);
            s.Check("the first slide is the one that was last but one or later",
                Order(after).Split(',')[^1] == was.Split(',')[0]);

            s.Check("moving a slide that is not there is refused",
                Slides.Move(file.Package, 99, 1) is Slides.Refused);
            s.Check("moving a slide to where it already is says so",
                Slides.Move(file.Package, 2, 2) is Slides.Done);
        }

        // ---- removing ----
        {
            var file = PlainFile.Open(Work());
            int count = file.Deck!.Slides.Count;
            var removed = Slides.Remove(file.Package, 2);
            s.Check("removing a slide says it did", removed is Slides.Done);

            var after = Reopen(file);
            s.Check("the deck still opens after a removal", after.Deck is not null);
            s.Check("there is one slide fewer", after.Deck!.Slides.Count == count - 1);
            s.Check("removing one that is not there is refused",
                Slides.Remove(file.Package, 99) is Slides.Refused);
        }

        // ---- the last slide cannot go ----
        {
            var file = PlainFile.Open(Work());
            while (file.Deck!.Slides.Count > 1)
            {
                if (Slides.Remove(file.Package, 1) is Slides.Refused) break;
                file = Reopen(file);
            }
            s.Check("a deck is never left with no slides at all",
                Slides.Remove(file.Package, 1) is Slides.Refused);
        }

        // ---- adding ----
        {
            var file = PlainFile.Open(Work());
            int count = file.Deck!.Slides.Count;
            var added = Slides.Add(file.Package, count);
            s.Check("adding a slide says it did", added is Slides.Done);

            var after = Reopen(file);
            s.Check("the deck still opens after an addition", after.Deck is not null);
            s.Check("there is one slide more", after.Deck!.Slides.Count == count + 1);

            // The new part has to be declared, related and reachable, or PowerPoint refuses the file.
            var types = System.Text.Encoding.UTF8.GetString(after.Package.Read("[Content_Types].xml"));
            s.Check("the new slide is declared in the content types", types.Contains("/ppt/slides/slide"));
            s.Check("the new slide has its own relationships part",
                after.Package.Parts.Any(p => p.Name.Contains("ppt/slides/_rels/")));

            var added2 = Slides.Add(after.Package, 0);
            s.Check("one can be added at the front too", added2 is Slides.Done);
            var after2 = Reopen(after);
            s.Check("and the deck still opens", after2.Deck is not null);
            s.Check("with two more than it started with", after2.Deck!.Slides.Count == count + 2);
        }

        // ---- what was already there is left alone ----
        {
            var file = PlainFile.Open(Work());
            int partsBefore = file.Package.Parts.Count;
            Slides.Move(file.Package, 1, 2);
            var after = Reopen(file);
            s.Check("moving a slide adds no parts to the file", after.Package.Parts.Count == partsBefore);
            s.Check("and the slides still have their words",
                after.Deck!.Slides.Any(x => x.Texts().Any(t => t.Lines.Any(l => l.Length > 0))));
        }

        return s;
    }
}
