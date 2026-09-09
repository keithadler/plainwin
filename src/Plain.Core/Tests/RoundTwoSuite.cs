namespace Plain.Core.Tests;

/// <summary>The features the second round of panels asked for, checked the same way as everything else.</summary>
public static class RoundTwoSuite
{
    public static Suite Run()
    {
        var s = new Suite("roundtwo");
        Counting(s);
        Comma(s);
        if (Fixtures.Missing(s, "doc.docx", "sheet.xlsx", "review.docx")) return s;
        Paper(s);
        Pictures(s);
        Settling(s);
        Lists(s);
        return s;
    }

    private static void Counting(Suite s)
    {
        s.Equal("nothing counts as nothing", 0, Counts.Of("").Words);
        s.Equal("one word", 1, Counts.Of("hello").Words);
        s.Equal("two words", 2, Counts.Of("hello there").Words);
        s.Equal("runs of spaces are not words", 2, Counts.Of("  hello    there  ").Words);
        s.Equal("lines are not words", 3, Counts.Of("one\ntwo\nthree").Words);
        s.Equal("punctuation travels with its word", 2, Counts.Of("hello, there!").Words);
        s.Equal("characters are counted", 5, Counts.Of("hello").Characters);
        s.Equal("and without spaces", 10, Counts.Of("hello there").CharactersWithoutSpaces);
        s.Equal("paragraphs are lines with something on them", 2, Counts.Of("one\n\ntwo\n").Paragraphs);
    }

    private static void Comma(Suite s)
    {
        s.Equal("a plain field needs no quotes", "abc", Csv.Quote("abc"));
        s.Equal("a comma forces quotes", "\"a,b\"", Csv.Quote("a,b"));
        s.Equal("a quote is doubled", "\"say \"\"hi\"\"\"", Csv.Quote("say \"hi\""));
        s.Equal("a leading space is protected", "\" a\"", Csv.Quote(" a"));

        var rows = Csv.Read("a,b,c\r\n1,2,3\r\n");
        s.Equal("two rows", 2, rows.Count);
        s.Equal("three fields", 3, rows[0].Count);
        s.Equal("read in order", "c", rows[0][2]);

        var quoted = Csv.Read("\"a,1\",\"say \"\"hi\"\"\",plain\n");
        s.Equal("a quoted comma stays in its field", "a,1", quoted[0][0]);
        s.Equal("a doubled quote becomes one", "say \"hi\"", quoted[0][1]);
        s.Equal("and the rest is read", "plain", quoted[0][2]);

        var multiline = Csv.Read("\"one\ntwo\",after\n");
        s.Equal("a newline inside quotes stays in the field", "one\ntwo", multiline[0][0]);
        s.Equal("one row, not two", 1, multiline.Count);

        var empty = Csv.Read("a,,c\n");
        s.Equal("an empty field is kept", "", empty[0][1]);
        s.Equal("so the columns still line up", 3, empty[0].Count);

        if (Fixtures.Missing(s, "sheet.xlsx")) return;
        var work = Fixtures.Copy("sheet.xlsx");
        try
        {
            var book = new Workbook(OpcPackage.Open(work));
            var text = Csv.Write(book.Sheets[0]);
            s.Check("a sheet writes out as a csv", text.StartsWith("Region,Q1"));
            s.Check("with a row per row", text.Split("\r\n", StringSplitOptions.RemoveEmptyEntries).Length >= 5);
            s.Check("a number comes out as a number, not as it is dressed", text.Contains("412800"));
            s.Check("so nothing needs quoting round it", !text.Contains("\"412,800\""));
            s.Check("and the screen is still available", Csv.Write(book.Sheets[0], ',', formatted: true).Contains("412,800"));

            int written = Csv.Into(book.Sheets[0], "name,amount\nwidget,12\ngadget,\"1,500\"\n", CellRef.Parse("H1"));
            s.Equal("every field lands", 6, written);
            s.Equal("text lands as text", "widget", book.Sheets[0].Read("H2").Display);
            s.Equal("a number lands as a number", "12", book.Sheets[0].Read("I2").Raw);
            s.Equal("a quoted comma stays one field", "1,500", book.Sheets[0].Read("I3").Display);

            // A csv must never be able to put a formula into somebody's workbook.
            Csv.Into(book.Sheets[0], "=1+1\n", CellRef.Parse("K1"));
            s.Check("a formula from a csv arrives as text", book.Sheets[0].Read("K1").Formula is null);
            s.Equal("and reads as what it said", "=1+1", book.Sheets[0].Read("K1").Display);
        }
        finally { try { File.Delete(work); } catch { } }
    }

    private static void Paper(Suite s)
    {
        foreach (var name in new[] { "doc.docx", "sheet.xlsx", "deck.pptx" })
        {
            var file = PlainFile.Open(Fixtures.Path_(name)!);
            var pdf = PdfExport.Build(file, name);
            s.Check($"{name} becomes a pdf", pdf.Bytes.Length > 400);
            s.Check($"{name}'s pdf starts as one should",
                    System.Text.Encoding.ASCII.GetString(pdf.Bytes, 0, 5) == "%PDF-");
            s.Check($"{name}'s pdf ends as one should",
                    System.Text.Encoding.ASCII.GetString(pdf.Bytes).TrimEnd().EndsWith("%%EOF"));
            s.Check($"{name}'s pdf has at least a page", pdf.Pages >= 1);
        }

        s.Check("a deck gets a page per slide", PdfExport.Build(PlainFile.Open(Fixtures.Path_("deck.pptx")!), "d").Pages == 2);
        s.Check("plain letters can be written", Pdf.CanWrite("Hello, world! 123 £5"));
        s.Check("and are not reported as missing", Pdf.Unwritable("Hello") == "");
        s.Check("characters outside the fonts are named", Pdf.Unwritable("hello 世界").Length == 2);
        s.Check("wrapping breaks a long line", Pdf.Wrap(new string('a', 20) + " " + new string('b', 20), 10, false, 60).Count > 1);
        s.Check("wrapping keeps a short line whole", Pdf.Wrap("short", 10, false, 400).Count == 1);
        s.Check("measuring grows with the text", Pdf.Measure("iiii", 10, false) < Pdf.Measure("MMMM", 10, false));
    }

    private static void Pictures(Suite s)
    {
        var plain = PlainFile.Open(Fixtures.Path_("doc.docx")!);
        s.Equal("a document with no pictures has none", 0, Media.In(plain).Count);
        s.Check("a png would be drawable", Media.Drawable("word/media/image1.png"));
        s.Check("a metafile would not", !Media.Drawable("word/media/image1.emf"));

        var corpus = Environment.GetEnvironmentVariable("PLAIN_CORPUS");
        if (corpus is null || !Directory.Exists(corpus)) { s.Check("no corpus to look for pictures in; skipped", true); return; }

        var withPictures = Directory.EnumerateFiles(corpus, "*.docx", SearchOption.AllDirectories)
            .Where(f => !Path.GetFileName(f).StartsWith("~$"))
            .Select(f => { try { return PlainFile.Open(f); } catch { return null; } })
            .FirstOrDefault(f => f is not null && Media.In(f).Count > 0);
        if (withPictures is null) { s.Check("no file in the corpus has a picture; skipped", true); return; }

        var pictures = Media.In(withPictures);
        s.Check("pictures are found", pictures.Count > 0);
        s.Check("each says what it is", pictures.All(p => p.Kind.Length > 0));
        s.Check("each has a readable size", pictures.All(p => p.Size.Length > 0));
        s.Check("and its bytes come back", pictures[0].Read().Length > 0);
        s.Equal("as many bytes as it said", pictures[0].Bytes, pictures[0].Read().Length);
    }

    private static void Settling(Suite s)
    {
        // Rejecting: what was struck out comes back, what was added goes.
        var work = Fixtures.Copy("review.docx");
        try
        {
            var file = PlainFile.Open(work);
            int rejected = Annotations.RejectRevisions(file);
            s.Equal("both changes were turned down", 2, rejected);
            file.Save();

            var after = PlainFile.Open(work);
            s.Equal("no tracked changes are left", 0, Annotations.Revisions(after).Count);
            var text = after.Document!.PlainText();
            s.Check("the struck out words came back", text.Contains("sixty (60) days"));
            s.Check("the added words went", !text.Contains("thirty (30) days"));
        }
        finally { try { File.Delete(work); } catch { } }

        // One at a time.
        var single = Fixtures.Copy("review.docx");
        try
        {
            var file = PlainFile.Open(single);
            var before = Annotations.Revisions(file);
            s.Equal("two to start with", 2, before.Count);

            s.Check("one can be settled on its own", Annotations.SettleRevision(file, 0, accept: true));
            file.Save();

            var after = PlainFile.Open(single);
            var left = Annotations.Revisions(after);
            s.Equal("one is left", 1, left.Count);
            s.Check("and it is the other one", left[0].Text != before[0].Text);
            s.Check("the accepted words are in the document", after.Document!.PlainText().Contains("thirty (30) days"));
        }
        finally { try { File.Delete(single); } catch { } }

        // One comment, not all of them.
        var comment = Fixtures.Copy("review.docx");
        try
        {
            var file = PlainFile.Open(comment);
            var before = Annotations.Comments(file);
            s.Equal("two comments to start with", 2, before.Count);
            s.Check("one can be removed on its own", Annotations.RemoveComment(file, 0));
            file.Save();

            var after = PlainFile.Open(comment);
            var left = Annotations.Comments(after);
            s.Equal("one comment is left", 1, left.Count);
            s.Equal("and it is the other one", before[1].Text, left[0].Text);
            s.Check("a comment that is not there is refused", !Annotations.RemoveComment(after, 9));
        }
        finally { try { File.Delete(comment); } catch { } }
    }

    private static void Lists(Suite s)
    {
        // A document Plain made carries list definitions, so it can take a list.
        var made = Path.Combine(Path.GetTempPath(), "plain-list-" + Guid.NewGuid().ToString("N") + ".docx");
        try
        {
            var file = PlainFile.Create(made);
            var doc = file.Document!;
            var (bullets, numbers) = doc.ListsAvailable();
            s.Check("a new document can do bullets", bullets);
            s.Check("and numbers", numbers);

            doc.SetText(0, "First point");
            s.Check("a block becomes a bullet", doc.SetList(0, bulleted: true, on: true));
            s.Check("and says it is one", doc.IsList(0));
            file.Save();

            var again = PlainFile.Open(made).Document!;
            s.Check("which survives a save", again.IsList(0));
            s.Equal("with its words", "First point", again.Read(0).Text);
            s.Equal("and reads as a list item", BlockKind.ListItem, again.Read(0).Kind);

            again.SetList(0, bulleted: true, on: false);
            s.Check("and can be taken out of the list", !again.IsList(0));
        }
        finally { try { File.Delete(made); } catch { } }

        // A document with no list definitions is told so rather than given a broken list.
        var plain = Fixtures.Copy("doc.docx");
        try
        {
            var doc = PlainFile.Open(plain).Document!;
            var (bullets, _) = doc.ListsAvailable();
            if (!bullets) s.Check("a document with no list definitions refuses one", !doc.SetList(0, true, true));
            else s.Check("this document already has list definitions", doc.SetList(0, true, true));
        }
        finally { try { File.Delete(plain); } catch { } }
    }
}
