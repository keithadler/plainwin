using System.Text;

namespace Plain.Core.Tests;

public static class OpcSuite
{
    public static Suite Run()
    {
        var s = new Suite("opc");

        // The CRC-32 check value from the standard's own test string.
        s.Equal("CRC-32 of \"123456789\"", 0xCBF43926u, Crc32.Of(Encoding.ASCII.GetBytes("123456789")));
        s.Equal("a leading slash is the same part", "word/document.xml", OpcPackage.Normalize("/word/document.xml"));

        s.Throws<OpcPackage.PackageException>("a file that is not a package is refused",
            () => OpcPackage.Read(Encoding.ASCII.GetBytes("this is plainly not a spreadsheet")));
        s.Throws<OpcPackage.PackageException>("an empty file is refused", () => OpcPackage.Read(Array.Empty<byte>()));

        if (Fixtures.Missing(s, "sheet.xlsx", "book.xlsx", "doc.docx", "deck.pptx")) return s;
        var path = Fixtures.Path_("sheet.xlsx")!;

        var original = File.ReadAllBytes(path);
        var pkg = OpcPackage.Open(path);
        s.Check("the package has parts", pkg.Parts.Count > 5);
        s.Check("every part decompresses", pkg.Parts.All(p => pkg.Read(p.Name).Length >= 0));
        s.Check("content types part is present", pkg.Has("[Content_Types].xml"));
        s.Check("content types part is XML", pkg.ReadText("[Content_Types].xml").TrimStart().StartsWith("<?xml"));
        s.Check("a missing part is reported, not invented", pkg.Find("xl/nothing.xml") is null);
        s.Throws<OpcPackage.PackageException>("reading a missing part throws", () => pkg.Read("xl/nothing.xml"));

        // The promise: saving a file nobody edited gives back the same bytes.
        s.Bytes("an unedited save is byte identical", original, pkg.ToBytes());
        s.Check("nothing is marked edited after only reading", pkg.Parts.All(p => !p.Edited));

        // The other half of the promise: editing one part leaves every other part alone.
        var edited = OpcPackage.Open(path);
        var before = edited.Parts.ToDictionary(p => p.Name, p => edited.Read(p.Name));
        edited.WriteText("docProps/core.xml", edited.ReadText("docProps/core.xml").Replace("</cp:coreProperties>", "</cp:coreProperties>"));
        edited.Write("xl/sharedStrings.xml", edited.Read("xl/sharedStrings.xml"));
        var saved = OpcPackage.Read(edited.ToBytes());
        s.Equal("the part count is unchanged by an edit", edited.Parts.Count, saved.Parts.Count);

        int identicalParts = 0;
        foreach (var p in saved.Parts)
            if (before.TryGetValue(p.Name, out var was) && was.AsSpan().SequenceEqual(saved.Read(p.Name))) identicalParts++;
        s.Equal("every part still holds the same content after a save", saved.Parts.Count, identicalParts);

        // A real edit must survive the round trip.
        var two = OpcPackage.Open(path);
        two.WriteText("docProps/app.xml", "<x>edited by a test</x>");
        var back = OpcPackage.Read(two.ToBytes());
        s.Equal("an edited part reads back changed", "<x>edited by a test</x>", back.ReadText("docProps/app.xml"));
        s.Bytes("a part next to the edit is untouched",
                before["xl/worksheets/sheet1.xml"], back.Read("xl/worksheets/sheet1.xml"));
        s.Check("the edited file is still a readable package", back.Parts.Count == pkg.Parts.Count);

        // Editing every part in turn must never corrupt the container.
        var all = OpcPackage.Open(path);
        foreach (var p in all.Parts.ToList()) all.Write(p.Name, all.Read(p.Name));
        var rewritten = OpcPackage.Read(all.ToBytes());
        s.Equal("rewriting every part keeps them all", pkg.Parts.Count, rewritten.Parts.Count);
        int same = rewritten.Parts.Count(p => before[p.Name].AsSpan().SequenceEqual(rewritten.Read(p.Name)));
        s.Equal("rewriting every part changes no content", rewritten.Parts.Count, same);

        // The preserved list must account for every part, once.
        var file = PlainFile.Open(path);
        var notes = file.Parts();
        s.Equal("every part appears in the preserved list", pkg.Parts.Count, notes.Count);
        s.Check("the sheet is listed as shown", notes.Any(n => n.Name.Contains("worksheets/") && n.Role == PartRole.Shown));
        s.Check("the theme is listed as preserved", notes.Any(n => n.Name.Contains("theme") && n.Role == PartRole.Preserved));
        s.Check("every part is described in words", notes.All(n => n.What.Length > 0));
        s.Check("no part is described as unknown twice over", notes.Count(n => n.What.StartsWith("Part of the file")) <= 1);
        var counts = file.Counts();
        s.Equal("nothing is edited before an edit", 0, counts.Edited);
        s.Equal("the counts add up", counts.Total, counts.Edited + counts.Kept);

        // The count a save reports must include edits still sitting in the model.
        var counted = PlainFile.Open(path);
        counted.Workbook!.Sheets[0].Set("A1", "counted");
        var after = counted.Counts();
        s.Check("an edit still in the model is counted before saving", after.Edited >= 1,
                $"reported {after.Edited} parts rewritten after changing a cell");
        s.Equal("the counts still add up after an edit", after.Total, after.Edited + after.Kept);

        // Finding text must reach every sheet, block and slide, and must look at formulas too.
        var book = new Workbook(OpcPackage.Open(Fixtures.Path_("book.xlsx")!));
        var cells = Search.InWorkbook(book, "seat");
        s.Equal("a search reaches every sheet", 2, cells.Count);          // "Seat" on Detail, "seats" in Notes
        s.Equal("a hit says where it is", "Detail!A2", cells[0].Where);
        s.Equal("hits come back in reading order", "Notes!A3", cells[1].Where);
        s.Check("search is not case sensitive", Search.InWorkbook(book, "LICENCES").Count == 1);
        s.Equal("a search for nothing finds nothing", 0, Search.InWorkbook(book, "").Count);
        s.Equal("a term that is not there finds nothing", 0, Search.InWorkbook(book, "zzz").Count);

        var sheetBook = new Workbook(OpcPackage.Open(path));
        s.Check("a formula is searchable by its function name", Search.InWorkbook(sheetBook, "sum").Count >= 1);

        var doc = new Document(OpcPackage.Open(Fixtures.Path_("doc.docx")!));
        var blocks = Search.InDocument(doc, "confidential");
        s.Check("a document search finds the clause", blocks.Count >= 1);
        s.Check("a document hit carries its text", blocks[0].Text.Length > 0);

        var deck = new Deck(OpcPackage.Open(Fixtures.Path_("deck.pptx")!));
        var slides = Search.InDeck(deck, "month-end");
        s.Equal("a deck search finds the line", 1, slides.Count);
        s.Equal("a deck hit says which slide", 2, slides[0].Slide);

        // Saving must leave the file it wrote to being the same file, not a new one wearing its name.
        var scratch = Fixtures.Copy("sheet.xlsx");
        try
        {
            var created = File.GetCreationTimeUtc(scratch);
            var wasOnDisk = File.ReadAllBytes(scratch);
            var again = PlainFile.Open(scratch);
            s.Check("a file just opened has not changed on disk", !again.ChangedOnDisk());
            s.Check("a normal file is not read only", !again.IsReadOnly());

            again.Workbook!.Sheets[0].Set("A1", "kept");
            again.Save();
            s.Check("the saved file is still there", File.Exists(scratch));
            s.Check("no temporary file is left behind", !File.Exists(scratch + ".plain-tmp"));
            // Replacing a file in place rather than swapping a new one in is a Windows guarantee; other systems
            // record creation time differently, so only Windows is held to it.
            if (OperatingSystem.IsWindows())
                s.Equal("the file keeps the moment it was created", created, File.GetCreationTimeUtc(scratch));
            else
                s.Check("creation time is only checked on Windows", true);
            s.Check("the contents did change", !File.ReadAllBytes(scratch).AsSpan().SequenceEqual(wasOnDisk));
            s.Check("saving updates what Plain remembers about the file", !again.ChangedOnDisk());

            // Something else writing to the file must be noticed.
            File.SetLastWriteTimeUtc(scratch, DateTime.UtcNow.AddMinutes(5));
            s.Check("a change made by something else is noticed", again.ChangedOnDisk());
        }
        finally { try { File.Delete(scratch); } catch { } }

        return s;
    }
}
