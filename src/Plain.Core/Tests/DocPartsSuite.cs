namespace Plain.Core.Tests;

/// <summary>
/// The parts of a document people forget are there: the lines along the top and bottom of every page, where a
/// paragraph sits across the page, and where its links actually go. And comparing two versions of a file.
/// </summary>
public static class DocPartsSuite
{
    private static string? Demo(string name)
    {
        var corpus = Environment.GetEnvironmentVariable("PLAIN_CORPUS");
        if (corpus is null) return null;
        var path = System.IO.Path.Combine(corpus, name);
        return File.Exists(path) ? path : null;
    }

    private static PlainFile Copy(string source)
    {
        var path = System.IO.Path.Combine(System.IO.Path.GetTempPath(),
            "plain-parts-" + Guid.NewGuid().ToString("N") + System.IO.Path.GetExtension(source));
        File.Copy(source, path, overwrite: true);
        return PlainFile.Open(path);
    }

    public static Suite Run()
    {
        var s = new Suite("docparts");

        // ---- headers and footers ----
        var review = Demo("review.docx");
        if (review is null)
            s.Check("no demo folder set (PLAIN_CORPUS); skipping headers and footers", true);
        else
        {
            var file = Copy(review);
            var bands = HeaderFooter.All(file);
            s.Check("the document's headers and footers are found", bands.Count >= 2);
            s.Check("one of them is a header", bands.Any(b => b.IsHeader));
            s.Check("one of them is a footer", bands.Any(b => !b.IsHeader));
            s.Check("the header says what it says",
                bands.Any(b => b.IsHeader && b.Text.Contains("in confidence")));
            s.Check("the footer too", bands.Any(b => !b.IsHeader && b.Text.Contains("annual review")));
            s.Check("each says which pages it is for", bands.All(b => b.Which.Length > 0));

            var header = bands.First(b => b.IsHeader);
            s.Check("a header can be changed",
                HeaderFooter.Write(file.Package, header.Part, "Pine Street Holdings"));
            s.Check("and says the new thing", HeaderFooter.Read(file.Package, header.Part) == "Pine Street Holdings");

            file.Flush();
            var back = PlainFile.Read(file.Package.ToBytes(), file.Path);
            s.Check("the document still opens", back.Document is not null);
            s.Check("and the change survived the save",
                HeaderFooter.All(back).Any(b => b.Text == "Pine Street Holdings"));
            s.Check("what was not touched is still there",
                HeaderFooter.All(back).Any(b => b.Text.Contains("annual review")));

            s.Check("spaces someone typed are kept",
                HeaderFooter.Write(file.Package, header.Part, "  spaced  ")
                && HeaderFooter.Read(file.Package, header.Part).StartsWith("  "));
            s.Check("a part that is not there is refused",
                !HeaderFooter.Write(file.Package, "word/nothing.xml", "x"));
        }

        // ---- where a paragraph sits ----
        {
            var path = System.IO.Path.Combine(System.IO.Path.GetTempPath(), "plain-jc-" + Guid.NewGuid().ToString("N") + ".docx");
            var file = PlainFile.Create(path);
            var doc = file.Document!;
            doc.SetText(0, "A line to move about the page.");

            s.Check("a paragraph starts saying nothing about where it sits", doc.ParagraphAlignment(0) == "general");
            s.Check("it can be centred", doc.SetParagraphAlignment(0, "centre") && doc.ParagraphAlignment(0) == "centre");
            s.Check("the American spelling works too", doc.SetParagraphAlignment(0, "center") && doc.ParagraphAlignment(0) == "centre");
            s.Check("it can be put right", doc.SetParagraphAlignment(0, "right") && doc.ParagraphAlignment(0) == "right");
            s.Check("and justified", doc.SetParagraphAlignment(0, "justify") && doc.ParagraphAlignment(0) == "justify");
            s.Check("and put back to whatever its style says",
                doc.SetParagraphAlignment(0, "general") && doc.ParagraphAlignment(0) == "general");
            s.Check("a word it does not know is refused", !doc.SetParagraphAlignment(0, "sideways"));
            s.Check("a paragraph that is not there is refused", !doc.SetParagraphAlignment(99, "left"));

            doc.SetParagraphAlignment(0, "centre");
            file.Flush();
            var back = PlainFile.Read(file.Package.ToBytes(), file.Path);
            s.Check("it survives a save", back.Document!.ParagraphAlignment(0) == "centre");
            s.Check("and the words are still there", back.Document.Read(0).Text.Contains("move about"));
        }

        // ---- links ----
        {
            s.Check("a file with no links reports none",
                Links.All(PlainFile.Create(System.IO.Path.Combine(System.IO.Path.GetTempPath(),
                    "plain-nolinks-" + Guid.NewGuid().ToString("N") + ".docx"))).Count == 0);

            s.Check("an ordinary web address is safe to open", Links.SafeToOpen("https://example.invalid/a"));
            s.Check("so is plain http", Links.SafeToOpen("http://example.invalid"));
            s.Check("and an email address", Links.SafeToOpen("mailto:someone@example.invalid"));
            s.Check("a file on disk is not offered", !Links.SafeToOpen("file:///C:/Windows/System32/cmd.exe"));
            s.Check("nor is anything that would run something", !Links.SafeToOpen("javascript:alert(1)"));
            s.Check("nor a Windows shell address", !Links.SafeToOpen("ms-msdt:/id"));
            s.Check("nor something that is not an address at all", !Links.SafeToOpen("just some words"));
            s.Check("nor an empty one", !Links.SafeToOpen(""));
        }

        // ---- comparing two files ----
        {
            var one = System.IO.Path.Combine(System.IO.Path.GetTempPath(), "plain-cmp-a-" + Guid.NewGuid().ToString("N") + ".xlsx");
            var before = PlainFile.Create(one);
            before.Workbook!.Sheets[0].Set("A1", "Licences");
            before.Workbook.Sheets[0].Set("B1", "100");
            before.Workbook.Sheets[0].Set("A2", "Hosting");
            before.Flush();
            var beforeBytes = before.Package.ToBytes();

            var after = PlainFile.Read(beforeBytes, one);
            s.Check("a file compared with itself has nothing changed",
                Compare.Between(PlainFile.Read(beforeBytes, one), after).Text.Count == 0);
            s.Check("and says how many parts were the same",
                Compare.Between(PlainFile.Read(beforeBytes, one), after).Same > 0);

            after.Workbook!.Sheets[0].Set("B1", "150");
            after.Workbook.Sheets[0].Set("A3", "Training");
            after.Flush();

            var report = Compare.Between(PlainFile.Read(beforeBytes, one), after);
            s.Check("a changed value is noticed", report.Text.Any(t => t.Text == "150" && t.Added));
            s.Check("and what it used to be", report.Text.Any(t => t.Text == "100" && !t.Added));
            s.Check("something new is noticed", report.Text.Any(t => t.Text == "Training" && t.Added));
            s.Check("something untouched is not mentioned", !report.Text.Any(t => t.Text == "Licences"));
            s.Check("it says which part changed", report.Parts.Any(p => p.How == "changed"));
            s.Check("and each change says where it was", report.Text.All(t => t.Where.Length > 0));

            var deckPath = System.IO.Path.Combine(System.IO.Path.GetTempPath(), "plain-cmp-c-" + Guid.NewGuid().ToString("N") + ".pptx");
            s.Check("two different kinds of file say so rather than pretending",
                Compare.Between(after, PlainFile.Create(deckPath)).Note is not null);
        }

        return s;
    }
}
