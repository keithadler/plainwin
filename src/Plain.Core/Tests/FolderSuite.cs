namespace Plain.Core.Tests;

/// <summary>
/// Searching a folder of files. The checks care that it finds what is there, that it never changes anything, and
/// that one unreadable file does not stop the rest: a folder search that gives up on the first bad file is a
/// folder search nobody can rely on.
/// </summary>
public static class FolderSuite
{
    public static Suite Run()
    {
        var s = new Suite("folder");
        if (Fixtures.Missing(s, "doc.docx", "sheet.xlsx", "deck.pptx")) return s;

        var dir = System.IO.Path.Combine(System.IO.Path.GetTempPath(), "plain-folder-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(dir);
        var inner = System.IO.Path.Combine(dir, "inside");
        Directory.CreateDirectory(inner);

        File.Copy(Fixtures.Path_("doc.docx")!, System.IO.Path.Combine(dir, "one.docx"));
        File.Copy(Fixtures.Path_("sheet.xlsx")!, System.IO.Path.Combine(dir, "two.xlsx"));
        File.Copy(Fixtures.Path_("deck.pptx")!, System.IO.Path.Combine(inner, "three.pptx"));

        // Things it should pass over: not an Office file, and Word's own leftovers.
        File.WriteAllText(System.IO.Path.Combine(dir, "notes.txt"), "Woodland Ave");
        File.WriteAllText(System.IO.Path.Combine(dir, "~$one.docx"), "not a document");
        // And something that looks like one and is not, so the search has to survive it.
        File.WriteAllText(System.IO.Path.Combine(dir, "broken.docx"), "this is not a package at all");

        var before = Directory.GetFiles(dir, "*", SearchOption.AllDirectories)
                              .ToDictionary(f => f, f => new FileInfo(f).Length);

        // ---- what it finds ----
        var wide = Folder.Search(dir, "the", deep: true);
        s.Check("it looked at the Office files", wide.Looked >= 3);
        s.Check("it did not look at the text file", !wide.Hits.Any(h => h.Path.EndsWith(".txt")));
        s.Check("it did not look at Word's leftover", !wide.Hits.Any(h => h.Path.Contains("~$")));
        s.Check("a file it could not read is reported, not thrown",
            wide.Troubles.Any(t => t.Path.EndsWith("broken.docx")));
        s.Check("and the reason is in words", wide.Troubles.All(t => t.Why.Length > 4));
        s.Check("one bad file does not stop the rest", wide.Hits.Count > 0);

        // ---- deep and shallow ----
        var shallow = Folder.Search(dir, "the", deep: false);
        s.Check("a shallow search stays in the folder it was given",
            !shallow.Hits.Any(h => h.Path.Contains("inside")));
        s.Check("a deep search goes into the folders inside it",
            wide.Looked > shallow.Looked);

        // ---- what it says about each hit ----
        foreach (var hit in wide.Hits)
        {
            s.Check($"{System.IO.Path.GetFileName(hit.Path)} says how many times", hit.Count > 0);
            s.Check($"{System.IO.Path.GetFileName(hit.Path)} shows where, without showing hundreds", hit.Places.Count is > 0 and <= 5);
        }

        // ---- nothing found, and nothing asked ----
        s.Check("something that is not there finds nothing",
            Folder.Search(dir, "zzqqxx-not-in-any-of-these", deep: true).Hits.Count == 0);
        s.Check("an empty term finds nothing rather than everything",
            Folder.Search(dir, "   ", deep: true).Hits.Count == 0);
        s.Check("a folder that is not there is not an error",
            Folder.Search(System.IO.Path.Combine(dir, "nowhere"), "the").Hits.Count == 0);

        // ---- the important one: searching changes nothing ----
        bool untouched = true;
        foreach (var (path, length) in before)
            if (!File.Exists(path) || new FileInfo(path).Length != length) { untouched = false; break; }
        s.Check("searching a folder leaves every file exactly as it was", untouched);

        // ---- it can be stopped and bounded ----
        s.Check("it stops when told to", Folder.Search(dir, "the", deep: true, cancelled: () => true).Looked == 0);
        s.Check("it looks at no more than it was allowed to",
            Folder.Search(dir, "the", deep: true, stopAfter: 1).Looked <= 1);

        try { Directory.Delete(dir, recursive: true); } catch { }
        return s;
    }
}
