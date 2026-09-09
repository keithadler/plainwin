namespace Plain.Core.Tests;

/// <summary>
/// Finding what travels with a file that you did not mean to send, and taking out what you choose.
///
/// The checks that matter most are the ones about restraint: that finding never changes anything, that removing
/// takes out only what was asked for, and that the things Plain will not remove stay put and stay listed. A tool
/// that quietly removed more than it said would be worse than one that removed nothing.
/// </summary>
public static class HiddenSuite
{
    public static Suite Run()
    {
        var s = new Suite("hidden");
        if (Fixtures.Missing(s, "review.docx", "doc.docx", "sheet.xlsx")) return s;

        PlainFile Work(string name)
        {
            var path = System.IO.Path.Combine(System.IO.Path.GetTempPath(),
                "plain-hidden-" + Guid.NewGuid().ToString("N") + System.IO.Path.GetExtension(name));
            File.Copy(Fixtures.Path_(name)!, path, overwrite: true);
            return PlainFile.Open(path);
        }

        // ---- a document with comments and tracked changes ----
        {
            var file = Work("review.docx");
            var found = Hidden.Find(file);
            s.Check("a marked up document has something to report", found.Count > 0);
            s.Check("the comments are reported", found.Any(f => f.Kind == Hidden.Comments));
            s.Check("the tracked changes are reported", found.Any(f => f.Kind == Hidden.Tracked));
            s.Check("each finding says how many", found.All(f => f.Count > 0));
            s.Check("and says it in words a person can read", found.All(f => f.What.Length > 4));

            // Looking must never change anything.
            var again = Hidden.Find(file);
            s.Check("looking twice reports the same thing", again.Count == found.Count);
            s.Check("looking does not change the file", !file.Package.Parts.Any(p => p.Edited));
        }

        // ---- removing only what was asked for ----
        {
            var file = Work("review.docx");
            int commentsBefore = Annotations.Comments(file).Count;
            int changesBefore = Annotations.Revisions(file).Count;
            s.Check("there were comments to begin with", commentsBefore > 0);
            s.Check("and tracked changes", changesBefore > 0);

            var did = Hidden.Remove(file, new[] { Hidden.Comments });
            s.Check("removing says what it did", did.Count > 0);
            s.Check("the comments are gone", Annotations.Comments(file).Count == 0);
            s.Check("the tracked changes were left alone", Annotations.Revisions(file).Count == changesBefore);

            var after = Hidden.Find(file);
            s.Check("and the report no longer mentions comments", !after.Any(f => f.Kind == Hidden.Comments));
            s.Check("but still mentions the tracked changes", after.Any(f => f.Kind == Hidden.Tracked));

            file.Flush();
            var back = PlainFile.Read(file.Package.ToBytes(), file.Path);
            s.Check("the document still opens afterwards", back.Document is not null);
        }

        // ---- names come out when asked ----
        {
            var file = Work("doc.docx");
            file.Properties.Set("Author", "Sam Rivera");
            file.Properties.Set("Last saved by", "Sam Rivera");
            file.Properties.Flush();

            var found = Hidden.Find(file);
            s.Check("a name in the properties is reported", found.Any(f => f.Kind == Hidden.People));
            s.Check("and the report names it", found.Any(f => f.What.Contains("Sam Rivera")));

            Hidden.Remove(file, new[] { Hidden.People });
            s.Check("the names come out", !Hidden.Find(file).Any(f => f.Kind == Hidden.People));
        }

        // ---- asking for nothing removes nothing ----
        {
            var file = Work("review.docx");
            int comments = Annotations.Comments(file).Count;
            var did = Hidden.Remove(file, Array.Empty<string>());
            s.Check("asking for nothing does nothing", did.Count == 0);
            s.Check("and leaves the comments alone", Annotations.Comments(file).Count == comments);
        }

        // ---- what Plain will not remove is still reported ----
        {
            var file = Work("review.docx");
            var found = Hidden.Find(file);
            foreach (var finding in found.Where(f => !f.CanRemove))
            {
                var did = Hidden.Remove(file, new[] { finding.Kind });
                s.Check($"asking to remove {finding.Kind} does not quietly do it", did.Count == 0);
            }
            s.Check("a plain file with nothing in it reports nothing at all",
                Hidden.Find(Work("sheet.xlsx")).All(f => f.Count > 0));
        }

        return s;
    }
}
