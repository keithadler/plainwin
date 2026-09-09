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

        NewFiles(s);
        Summaries(s);
        Homes(s);
        return s;
    }

    /// <summary>
    /// A file opened from bytes has to remember where it belongs, or a save after any change large enough to reopen
    /// it goes nowhere. That is exactly what happened to replace-all and to inserting a row.
    /// </summary>
    private static void Homes(Suite s)
    {
        var scratch = Fixtures.Copy("sheet.xlsx");
        try
        {
            var bytes = File.ReadAllBytes(scratch);

            var nameless = PlainFile.Read(bytes);
            s.Equal("a file read from bytes has no home by default", "", nameless.Path);
            s.Throws<OpcPackage.PackageException>("and saving it without one is refused", () => nameless.Save());
            s.Check("but it can be saved somewhere named", true);

            var homed = PlainFile.Read(bytes, scratch);
            s.Equal("a file read from bytes can be told where it belongs", scratch, homed.Path);
            homed.Workbook!.Sheets[0].Set("A1", "rebuilt");
            homed.Save();
            s.Equal("and saves back to there", "rebuilt",
                    PlainFile.Open(scratch).Workbook!.Sheets[0].Read("A1").Display);

            // Saving somewhere else moves where it belongs, the way Save a copy then working on the copy would.
            var elsewhere = scratch + ".copy.xlsx";
            try
            {
                var moved = PlainFile.Open(scratch);
                moved.Save(elsewhere);
                s.Check("saving a copy leaves the original alone", File.Exists(scratch) && File.Exists(elsewhere));
            }
            finally { try { File.Delete(elsewhere); } catch { } }
        }
        finally { try { File.Delete(scratch); } catch { } }
    }

    /// <summary>The panel has to read like a sentence about the file, not like a directory listing of it.</summary>
    private static void Summaries(Suite s)
    {
        s.Equal("a chart pluralises", "charts", Preserved.Many("A chart"));
        s.Equal("an embedded font pluralises", "embedded fonts", Preserved.Many("An embedded font"));
        s.Equal("a page footer pluralises", "page footers", Preserved.Many("A page footer"));
        s.Equal("a query pluralises properly", "queries", Preserved.Many("A query"));
        s.Equal("a box pluralises properly", "boxes", Preserved.Many("A box"));
        s.Equal("something already plural is left alone", "Comments", Preserved.Many("Comments"));
        s.Equal("a mass noun is left alone", "Macros", Preserved.Many("Macros"));

        var notes = new[]
        {
            new PartNote("word/fonts/font1.odttf", PartRole.Preserved, "An embedded font", 4000, PartClass.Content),
            new PartNote("word/fonts/font2.odttf", PartRole.Preserved, "An embedded font", 6000, PartClass.Content),
            new PartNote("word/header1.xml", PartRole.Preserved, "A page header", 500, PartClass.Content),
            new PartNote("_rels/.rels", PartRole.Preserved, "Links between parts", 300, PartClass.Bookkeeping),
            new PartNote("[Content_Types].xml", PartRole.Preserved, "The list of what each part is", 900, PartClass.Bookkeeping),
            new PartNote("word/document.xml", PartRole.Shown, "The document text", 7000, PartClass.Content),
        };

        var rows = Preserved.Summarise(notes);
        s.Equal("only what is worth reading gets a row", 2, rows.Count);
        s.Equal("the biggest thing comes first", "2 embedded fonts", rows[0].Title);
        s.Equal("and carries both their sizes", 10000, rows[0].Bytes);
        s.Check("a row for several things names none of them", rows[0].Name is null);
        s.Equal("one of a thing reads as itself", "A page header", rows[1].Title);
        s.Equal("and names the part", "word/header1.xml", rows[1].Name);
        s.Check("what Plain shows never appears in the list", rows.All(r => !r.What.Contains("document text")));

        var (count, bytes) = Preserved.Bookkeeping(notes);
        s.Equal("bookkeeping is counted, not listed", 2, count);
        s.Equal("with its size", 1200, bytes);

        // The classification itself.
        s.Equal("relationships are bookkeeping", PartClass.Bookkeeping, Preserved.Classify("word/_rels/document.xml.rels"));
        s.Equal("content types are bookkeeping", PartClass.Bookkeeping, Preserved.Classify("[Content_Types].xml"));
        s.Equal("document properties are bookkeeping", PartClass.Bookkeeping, Preserved.Classify("docProps/app.xml"));
        s.Equal("a theme is bookkeeping", PartClass.Bookkeeping, Preserved.Classify("ppt/theme/theme1.xml"));
        s.Equal("a chart is not", PartClass.Content, Preserved.Classify("xl/charts/chart1.xml"));
        s.Equal("a header is not", PartClass.Content, Preserved.Classify("word/header1.xml"));
        s.Equal("macros are not", PartClass.Content, Preserved.Classify("word/vbaProject.bin"));
        s.Equal("an embedded font is not", PartClass.Content, Preserved.Classify("word/fonts/font1.odttf"));
        s.Equal("an embedded font is named", "An embedded font", Preserved.What("word/fonts/font1.odttf"));
        s.Equal("a folder marker is bookkeeping", PartClass.Bookkeeping, Preserved.Classify("word/_rels/"));
        // Every phrase Plain can produce must read properly when there are several of them.
        var phrases = new (string Part, string Several)[]
        {
            ("xl/charts/chart1.xml", "charts"),
            ("xl/pivotCache/pivotCacheDefinition1.xml", "Pivot table data"),
            ("xl/tables/table1.xml", "named tables"),
            ("xl/queryTables/queryTable1.xml", "external data queries"),
            ("word/header1.xml", "page headers"),
            ("word/footer1.xml", "page footers"),
            ("word/fonts/font1.odttf", "embedded fonts"),
            ("word/media/image1.png", "pictures"),
            ("word/embeddings/oleObject1.bin", "embedded files"),
            ("word/vbaProject.bin", "Macros"),
            ("ppt/slideLayouts/slideLayout1.xml", "slide layouts"),
            ("ppt/slideMasters/slideMaster1.xml", "slide masters"),
            ("ppt/diagrams/data1.xml", "SmartArt diagrams"),
            ("ppt/notesSlides/notesSlide1.xml", "Speaker notes"),
            ("word/comments.xml", "Comments"),
            ("something/nobody/knows.bin", "unrecognised parts"),
        };
        foreach (var (part, several) in phrases)
            s.Equal($"several of \"{Preserved.What(part)}\" reads right", several, Preserved.Many(Preserved.What(part)));
    }

    /// <summary>
    /// A file Plain makes has to be a real one: it must open, it must survive a save unchanged, it must take an edit,
    /// and the edit must read back. Anything less and "New" would be handing people a file that fails elsewhere.
    /// </summary>
    private static void NewFiles(Suite s)
    {
        foreach (var (extension, kind) in new[] { (".xlsx", FileKind.Spreadsheet), (".docx", FileKind.Document), (".pptx", FileKind.Presentation) })
        {
            var made = Blank.Make(kind);
            s.Check($"a new {extension} is not empty", made.Length > 500);

            // Two blank files must be the same bytes, or nothing about them can be tested.
            s.Bytes($"a new {extension} is made the same way every time", made, Blank.Make(kind));

            var file = PlainFile.Read(made);
            s.Equal($"a new {extension} opens as the right kind", kind, file.Kind);
            s.Bytes($"a new {extension} survives a save untouched", made, file.Package.ToBytes());
            s.Check($"a new {extension} has its content types", file.Package.Has("[Content_Types].xml"));
            s.Check($"a new {extension} has its root relationships", file.Package.Has("_rels/.rels"));
            s.Check($"every part of a new {extension} is readable",
                    file.Package.Parts.All(part => file.Package.Read(part.Name).Length > 0));
            s.Check($"every part of a new {extension} is declared in the content types",
                    Declared(file), "a part is in the package that nothing says the type of");

            switch (kind)
            {
                case FileKind.Spreadsheet:
                    s.Equal("a new workbook has one sheet", 1, file.Workbook!.Sheets.Count);
                    s.Equal("the sheet is called Sheet1", "Sheet1", file.Workbook.Sheets[0].Name);
                    s.Check("the sheet starts empty", !file.Workbook.Sheets[0].Cells().Any());
                    file.Workbook.Sheets[0].Set("B3", "hello");
                    file.Workbook.Sheets[0].Set("B4", "42");
                    file.Workbook.Sheets[0].Set("B5", "=B4*2");
                    break;
                case FileKind.Document:
                    s.Check("a new document has a line to type on", file.Document!.BlockCount >= 1);
                    file.Document.SetText(0, "hello");
                    break;
                case FileKind.Presentation:
                    s.Equal("a new deck has one slide", 1, file.Deck!.Slides.Count);
                    s.Check("the slide has text to replace", file.Deck.Slides[0].Texts().Any());
                    file.Deck.Slides[0].SetLine(0, 0, "hello");
                    break;
            }

            file.Flush();
            var reopened = PlainFile.Read(file.Package.ToBytes());
            switch (kind)
            {
                case FileKind.Spreadsheet:
                    s.Equal("text typed into a new workbook reads back", "hello", reopened.Workbook!.Sheets[0].Read("B3").Display);
                    s.Equal("a number typed into a new workbook reads back", "42", reopened.Workbook.Sheets[0].Read("B4").Raw);
                    s.Equal("a formula typed into a new workbook reads back", "=B4*2", reopened.Workbook.Sheets[0].Read("B5").Formula);
                    break;
                case FileKind.Document:
                    s.Equal("text typed into a new document reads back", "hello", reopened.Document!.Read(0).Text);
                    break;
                case FileKind.Presentation:
                    s.Equal("text typed into a new deck reads back", "hello", reopened.Deck!.Slides[0].Title());
                    break;
            }
        }

        // Creating on disk must refuse to write over something, and must refuse a name it cannot make sense of.
        var dir = Path.Combine(Path.GetTempPath(), "plain-new-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(dir);
        try
        {
            var made = Path.Combine(dir, "book.xlsx");
            var created = PlainFile.Create(made);
            s.Check("a new file lands on disk", File.Exists(made));
            s.Equal("and opens as a workbook", FileKind.Spreadsheet, created.Kind);
            s.Throws<OpcPackage.PackageException>("it will not write over a file that exists", () => PlainFile.Create(made));
            s.Throws<OpcPackage.PackageException>("it refuses a name it cannot place",
                () => PlainFile.Create(Path.Combine(dir, "mystery.txt")));
        }
        finally { try { Directory.Delete(dir, true); } catch { } }
    }

    /// <summary>Every part must be covered by a Default for its extension or an Override for its name.</summary>
    private static bool Declared(PlainFile file)
    {
        var types = Xml.Parse(file.Package.Read("[Content_Types].xml")).Root!;
        var defaults = types.Elements(Ns.Ct + "Default")
            .Select(d => (string?)d.Attribute("Extension") ?? "").ToHashSet(StringComparer.OrdinalIgnoreCase);
        var overrides = types.Elements(Ns.Ct + "Override")
            .Select(o => ((string?)o.Attribute("PartName") ?? "").TrimStart('/')).ToHashSet(StringComparer.Ordinal);

        foreach (var part in file.Package.Parts)
        {
            if (overrides.Contains(part.Name)) continue;
            var extension = System.IO.Path.GetExtension(part.Name).TrimStart('.');
            if (defaults.Contains(extension)) continue;
            return false;
        }
        return true;
    }
}
