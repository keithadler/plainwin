namespace Plain.Core.Tests;

public static class FormatSuite
{
    public static Suite Run()
    {
        var s = new Suite("format");
        if (Fixtures.Missing(s, "doc.docx", "sheet.xlsx")) return s;

        // A document: bold, italic and headings.
        var work = Fixtures.Copy("doc.docx");
        try
        {
            var file = PlainFile.Open(work);
            var doc = file.Document!;
            int target = doc.Blocks().First(b => b.Text.Contains("thirty (30) days")).Index;
            int untouched = doc.Blocks().First(b => b.Text.Contains("Discovery and audit")).Index;
            string text = doc.Read(target).Text;

            s.Check("it does not start bold", !doc.IsAll(target, "b"));
            doc.SetMark(target, "b", true);
            s.Check("it is bold now", doc.IsAll(target, "b"));
            s.Equal("and its words are unchanged", text, doc.Read(target).Text);

            doc.SetMark(target, "i", true);
            s.Check("and italic as well", doc.IsAll(target, "i"));
            s.Check("still bold", doc.IsAll(target, "b"));

            file.Save();
            var again = PlainFile.Open(work).Document!;
            s.Check("bold survives a save", again.IsAll(target, "b"));
            s.Check("italic survives a save", again.IsAll(target, "i"));
            s.Equal("the words survive too", text, again.Read(target).Text);
            s.Check("a neighbouring block is not bold", !again.IsAll(untouched, "b"));

            again.SetMark(target, "b", false);
            s.Check("bold comes off again", !again.IsAll(target, "b"));
            s.Check("italic stays on", again.IsAll(target, "i"));
        }
        finally { try { File.Delete(work); } catch { } }

        // Headings, and only the ones the document has.
        var headings = Fixtures.Copy("doc.docx");
        try
        {
            var file = PlainFile.Open(headings);
            var doc = file.Document!;
            var kinds = doc.AvailableKinds();
            s.Check("ordinary text is always available", kinds.Contains(BlockKind.Paragraph));

            int target = doc.Blocks().First(b => b.Text.Contains("thirty (30) days")).Index;
            if (kinds.Contains(BlockKind.Heading2))
            {
                s.Check("a heading can be applied", doc.SetKind(target, BlockKind.Heading2));
                s.Equal("and the block reads as one", BlockKind.Heading2, doc.Read(target).Kind);
                file.Save();
                s.Equal("which survives a save", BlockKind.Heading2,
                        PlainFile.Open(headings).Document!.Read(target).Kind);

                doc = PlainFile.Open(headings).Document!;
                doc.SetKind(target, BlockKind.Paragraph);
                s.Equal("and can be taken off again", BlockKind.Paragraph, doc.Read(target).Kind);
            }
            else s.Check("this document has no heading styles to apply", true);
        }
        finally { try { File.Delete(headings); } catch { } }

        // A new document must be able to take a heading, since Plain writes the styles into it.
        var made = Path.Combine(Path.GetTempPath(), "plain-fmt-" + Guid.NewGuid().ToString("N") + ".docx");
        try
        {
            var file = PlainFile.Create(made);
            var doc = file.Document!;
            s.Check("a new document offers headings", doc.AvailableKinds().Contains(BlockKind.Heading1));
            doc.SetText(0, "A Title");
            s.Check("and takes one", doc.SetKind(0, BlockKind.Heading1));
            file.Save();
            var again = PlainFile.Open(made).Document!;
            s.Equal("which survives a save", BlockKind.Heading1, again.Read(0).Kind);
            s.Equal("with its words", "A Title", again.Read(0).Text);
        }
        finally { try { File.Delete(made); } catch { } }

        // A spreadsheet: bold cells.
        var book = Fixtures.Copy("sheet.xlsx");
        try
        {
            var file = PlainFile.Open(book);
            var sheet = file.Workbook!.Sheets[0];
            var headers = Enumerable.Range(1, 5).Select(c => new CellRef(c, 1)).ToList();
            string value = sheet.Read("A1").Display;

            s.Check("the header row is not bold to start", !sheet.HasWeight(headers[0], "b"));
            sheet.SetWeight(headers, bold: true, italic: null);
            s.Check("the header row is bold now", headers.All(c => sheet.HasWeight(c, "b")));
            s.Equal("and its text is unchanged", value, sheet.Read("A1").Display);
            s.Check("a cell below is untouched", !sheet.HasWeight(new CellRef(1, 2), "b"));

            file.Save();
            var again = PlainFile.Open(book).Workbook!.Sheets[0];
            s.Check("bold survives a save", again.HasWeight(new CellRef(1, 1), "b"));
            s.Equal("and so does the value", value, again.Read("A1").Display);

            // Bold plus a number format must both stick.
            again.SetFormat(new[] { new CellRef(2, 2) }, "0.0%");
            again.SetWeight(new[] { new CellRef(2, 2) }, bold: true, italic: null);
            s.Equal("a cell keeps its format when it goes bold", "0.0%", again.FormatOf(new CellRef(2, 2)));
            s.Check("and is bold", again.HasWeight(new CellRef(2, 2), "b"));
        }
        finally { try { File.Delete(book); } catch { } }

        return s;
    }
}
