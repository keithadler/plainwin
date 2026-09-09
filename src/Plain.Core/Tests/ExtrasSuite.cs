namespace Plain.Core.Tests;

/// <summary>
/// Borders, tidying a list, tracing a formula, the choices a cell allows, describing a picture, speaker notes,
/// making a link, putting a picture in, and offering what is already in the column.
///
/// The checks that matter most are the refusals and the not-writing-over: splitting a column must never destroy
/// the one beside it, and Plain must not write a link it would refuse to open.
/// </summary>
public static class ExtrasSuite
{
    private static PlainFile Fresh(string extension = ".xlsx")
    {
        var path = System.IO.Path.Combine(System.IO.Path.GetTempPath(),
            "plain-extra-" + Guid.NewGuid().ToString("N") + extension);
        return PlainFile.Create(path);
    }

    private static PlainFile Reopen(PlainFile file)
    {
        file.Flush();
        return PlainFile.Read(file.Package.ToBytes(), file.Path);
    }

    private static string? Demo(string name)
    {
        var corpus = Environment.GetEnvironmentVariable("PLAIN_CORPUS");
        if (corpus is null) return null;
        var path = System.IO.Path.Combine(corpus, name);
        return File.Exists(path) ? path : null;
    }

    private static PlainFile CopyOf(string source)
    {
        var path = System.IO.Path.Combine(System.IO.Path.GetTempPath(),
            "plain-extra-" + Guid.NewGuid().ToString("N") + System.IO.Path.GetExtension(source));
        File.Copy(source, path, overwrite: true);
        return PlainFile.Open(path);
    }

    public static Suite Run()
    {
        var s = new Suite("extras");

        // ---- lines round a cell ----
        {
            var file = Fresh();
            var sheet = file.Workbook!.Sheets[0];
            sheet.Set("A1", "boxed");

            s.Check("a cell starts with no lines round it", sheet.BorderAt(CellRef.Parse("A1")).Count == 0);
            sheet.SetBorder(new[] { CellRef.Parse("A1") }, new[] { "all" }, "thin", "808080");
            var sides = sheet.BorderAt(CellRef.Parse("A1"));
            s.Check("all four sides get a line", sides.Count == 4);
            s.Check("and they are the weight asked for", sides.All(x => x.Style == "thin"));

            sheet.SetBorder(new[] { CellRef.Parse("A1") }, new[] { "bottom" }, "thick", "");
            var after = sheet.BorderAt(CellRef.Parse("A1"));
            s.Check("one side can be changed on its own", after.Any(x => x.Side == "bottom" && x.Style == "thick"));
            s.Check("and the others keep what they had", after.Any(x => x.Side == "top" && x.Style == "thin"));

            sheet.SetBorder(new[] { CellRef.Parse("A1") }, new[] { "all" }, "", "");
            s.Check("and they can all be taken away", sheet.BorderAt(CellRef.Parse("A1")).Count == 0);

            sheet.SetBorder(new[] { CellRef.Parse("A1") }, new[] { "all" }, "medium", "C00000");
            var back = Reopen(file);
            s.Check("lines survive a save", back.Workbook!.Sheets[0].BorderAt(CellRef.Parse("A1")).Count == 4);
        }

        // ---- taking out rows that say the same thing ----
        {
            var file = Fresh();
            var book = file.Workbook!;
            var sheet = book.Sheets[0];
            foreach (var (r, name) in new[] { (1, "Pear"), (2, "Apple"), (3, "Pear"), (4, "Cherry"), (5, "apple") })
                sheet.Set(new CellRef(1, r), name);

            var done = Tidy.RemoveDuplicates(book, sheet, 1, 5, 1, 1);
            s.Check("duplicates are taken out", done is Tidy.Done { Rows: 2 });
            s.Check("the first of each stays", sheet.Read("A1").Display == "Pear");
            s.Check("the order of what is left is kept", sheet.Read("A2").Display == "Apple");
            s.Check("and it does not mind the case", sheet.Read("A3").Display == "Cherry");
            s.Check("what was pushed up leaves empty cells behind", sheet.Read("A4").Kind == CellKind.Empty);

            var none = Tidy.RemoveDuplicates(book, sheet, 1, 3, 1, 1);
            s.Check("a list with no duplicates says so", none is Tidy.Done { Rows: 0 });

            sheet.Set("B1", "=A1");
            s.Check("a formula in the block is refused",
                Tidy.RemoveDuplicates(book, sheet, 1, 3, 1, 2) is Tidy.Refused);
        }

        // ---- splitting a column ----
        {
            var file = Fresh();
            var book = file.Workbook!;
            var sheet = book.Sheets[0];
            sheet.Set("A1", "Rivera, Sam");
            sheet.Set("A2", "Adler, Keith");
            sheet.Set("A3", "nobody");

            var done = Tidy.SplitColumn(book, sheet, 1, 1, 3, ",");
            s.Check("a column splits", done is Tidy.Done { Rows: 2 });
            s.Check("the first piece stays where it was", sheet.Read("A1").Display == "Rivera");
            s.Check("the second goes to the right", sheet.Read("B1").Display == "Sam");
            s.Check("spaces round a piece are trimmed", sheet.Read("B2").Display == "Keith");
            s.Check("a row with nothing to split is left alone", sheet.Read("A3").Display == "nobody");

            // The one that matters: it must never write over the column beside it.
            var file2 = Fresh();
            var book2 = file2.Workbook!;
            var sheet2 = book2.Sheets[0];
            sheet2.Set("A1", "one,two");
            sheet2.Set("B1", "do not lose me");
            var refused = Tidy.SplitColumn(book2, sheet2, 1, 1, 1, ",");
            s.Check("splitting is refused when it would write over something", refused is Tidy.Refused);
            s.Check("and it says which column", refused is Tidy.Refused r && r.Reason.Contains("B"));
            s.Check("and nothing was touched", sheet2.Read("B1").Display == "do not lose me");
            s.Check("nothing at all in the column says so",
                Tidy.SplitColumn(book2, sheet2, 3, 1, 1, ",") is Tidy.Done { Rows: 0 });
            s.Check("no separator is refused", Tidy.SplitColumn(book2, sheet2, 1, 1, 1, "") is Tidy.Refused);
        }

        // ---- what a formula reads, and what reads a cell ----
        {
            var file = Fresh();
            var book = file.Workbook!;
            var sheet = book.Sheets[0];
            sheet.Set("A1", "10");
            sheet.Set("A2", "20");
            sheet.Set("B1", "=SUM(A1:A2)");
            sheet.Set("C1", "=B1*2");

            var reads = Traces.Reads(book, sheet, CellRef.Parse("B1"));
            s.Check("a formula says what it reads", reads.Count == 1);
            s.Check("and names the range", reads.Count > 0 && reads[0].Where.Contains("A1:A2"));
            s.Check("and says what is in there", reads.Count > 0 && reads[0].What.Contains("2"));
            s.Check("a cell with no formula reads nothing",
                Traces.Reads(book, sheet, CellRef.Parse("A1")).Count == 0);

            var readBy = Traces.ReadBy(book, sheet, CellRef.Parse("A1"));
            s.Check("a cell says what reads it", readBy.Count == 1);
            s.Check("and names the formula", readBy.Count > 0 && readBy[0].What.Contains("SUM"));
            s.Check("a cell nothing reads says so",
                Traces.ReadBy(book, sheet, CellRef.Parse("C1")).Count == 0);
            s.Check("a cell read through a chain names only what reads it directly",
                Traces.ReadBy(book, sheet, CellRef.Parse("B1")).Count == 1);
        }

        // ---- the choices a cell allows ----
        var quarter = Demo("quarter.xlsx");
        if (quarter is null)
            s.Check("no demo folder set (PLAIN_CORPUS); skipping the choices checks", true);
        else
        {
            var file = CopyOf(quarter);
            var detail = file.Workbook!.Sheets.First(x => x.Name == "Detail");
            var rules = Choices.On(file.Package, detail);
            s.Check("a rule on a sheet is found", rules.Count >= 1);
            s.Check("it says what kind it is", rules.Count > 0 && rules[0].Kind.Length > 0);
            s.Check("a list rule says what the choices are", rules.Any(r => r.Allowed.Count >= 3));
            s.Check("and one of them is the one we put there",
                rules.Any(r => r.Allowed.Contains("Overdue")));
            s.Check("a cell inside it finds its rule", Choices.At(file.Package, detail, CellRef.Parse("E2")) is not null);
            s.Check("a cell outside it finds none", Choices.At(file.Package, detail, CellRef.Parse("A1")) is null);
        }

        // ---- speaker notes ----
        var deck = Demo("woodland.pptx");
        if (deck is null)
            s.Check("no demo deck; skipping the speaker notes checks", true);
        else
        {
            var file = CopyOf(deck);
            var notes = Described.Notes(file);
            s.Check("the deck's speaker notes are found", notes.Count >= 3);
            s.Check("they say something", notes.Any(n => n.Text.Contains("twenty minutes")));
            s.Check("each knows which slide it is for", notes.All(n => n.Slide >= 1));

            s.Check("notes can be changed",
                notes.Count > 0 && Described.WriteNotes(file.Package, notes[0].Part, "Say hello. Then get on with it."));
            var after = Described.Notes(Reopen(file));
            s.Check("and the change survives a save", after.Any(n => n.Text.Contains("get on with it")));
            s.Check("the others are untouched", after.Any(n => n.Text.Contains("twenty per cent")
                                                            || n.Text.Contains("March")));
        }

        // ---- links written into a document ----
        {
            var file = Fresh(".docx");
            file.Document!.SetText(0, "Our terms are on the website.");

            s.Check("something that would run a program is refused",
                Insert.Link(file, 0, "javascript:alert(1)") is Insert.Refused);
            s.Check("a file on disk is refused too",
                Insert.Link(file, 0, "file:///C:/Windows/System32/cmd.exe") is Insert.Refused);
            s.Check("and the document has no link on it after those",
                Links.All(file).Count == 0);

            var written = Insert.Link(file, 0, "https://example.invalid/terms");
            s.Check("an ordinary address is written", written is Insert.Done,
                written is Insert.Refused why ? why.Reason : "");
            var back = file.Reopened();
            s.Check("the document still opens", back.Document is not null);
            var links = Links.All(back);
            s.Check("and the link is there", links.Count == 1, $"found {links.Count}");
            s.Check("with the words it was put on", links.Count > 0 && links[0].Text.Contains("Our terms"));
            s.Check("and where it goes", links.Count > 0 && links[0].Target.Contains("example.invalid"));

            s.Check("a second link on the same paragraph is refused",
                Insert.Link(back, 0, "https://example.invalid/other") is Insert.Refused);
            s.Check("a link can be taken off", Insert.Unlink(back, 0) is Insert.Done);
            var plain = back.Reopened();
            s.Check("and then there is none", Links.All(plain).Count == 0);
            s.Check("but the words are still there", plain.Document!.Read(0).Text.Contains("Our terms"));
        }

        // ---- a picture put into a document ----
        {
            var file = Fresh(".docx");
            file.Document!.SetText(0, "A document with a picture in it.");

            var png = System.IO.Path.Combine(System.IO.Path.GetTempPath(), "plain-dot-" + Guid.NewGuid().ToString("N") + ".png");
            // The smallest PNG there is: one pixel.
            File.WriteAllBytes(png, Convert.FromBase64String(
                "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg=="));

            s.Check("a file that is not there is refused", Insert.Picture(file, png + ".missing", 5) is Insert.Refused);
            s.Check("a kind Plain does not put in is refused",
                Insert.Picture(file, System.IO.Path.ChangeExtension(png, ".tiff"), 5) is Insert.Refused);

            s.Check("a picture goes in", Insert.Picture(file, png, 6) is Insert.Done);
            var back = file.Reopened();
            s.Check("the document still opens", back.Document is not null);
            s.Check("the picture is in the file", Media.In(back).Count >= 1);
            s.Check("and the words are still there", back.Document!.Read(0).Text.Contains("picture in it"));
            s.Check("it is described in the content types",
                System.Text.Encoding.UTF8.GetString(back.Package.Read("[Content_Types].xml")).Contains("image/png"));

            var described = Described.Pictures(back);
            s.Check("the picture is found to describe", described.Count >= 1, $"found {described.Count}");
            s.Check("and reported once, not once per element it is named on", described.Count == 1,
                $"found {described.Count}");
            s.Check("and starts with no description", described.All(x => x.Text.Length == 0));
            if (described.Count > 0)
            {
                s.Check("a description can be given",
                    Described.DescribePictures(back, described[0].Where, described[0].Name, "A single dot.") >= 1);
                s.Check("and it survives a save",
                    Described.Pictures(back.Reopened()).Any(x => x.Text == "A single dot."));
            }
        }

        // ---- offering what is already in the column ----
        {
            var file = Fresh();
            var sheet = file.Workbook!.Sheets[0];
            sheet.Set("A1", "Woodland Ave Partners");
            sheet.Set("A2", "Harbour Lane Clinic");
            sheet.Set("A3", "12345");

            s.Check("it offers what is above", sheet.Suggest(1, 4, "Wood") == "Woodland Ave Partners");
            s.Check("it does not mind the case", sheet.Suggest(1, 4, "wood") == "Woodland Ave Partners");
            s.Check("it offers nothing when nothing matches", sheet.Suggest(1, 4, "Zebra") is null);
            s.Check("it offers nothing for an empty start", sheet.Suggest(1, 4, "") is null);
            s.Check("it does not offer to finish a number", sheet.Suggest(1, 4, "123") is null);
            s.Check("it does not offer what is below", sheet.Suggest(1, 1, "Wood") is null);
            s.Check("it does not offer the same thing back",
                sheet.Suggest(1, 4, "Woodland Ave Partners") is null);
        }

        return s;
    }
}
