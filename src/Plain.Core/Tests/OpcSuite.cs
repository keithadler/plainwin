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

        var path = Fixtures.Path_("sheet.xlsx");
        if (path is null) { s.Check("fixtures found", false, "tests/fixtures/sheet.xlsx is missing"); return s; }

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

        return s;
    }
}
