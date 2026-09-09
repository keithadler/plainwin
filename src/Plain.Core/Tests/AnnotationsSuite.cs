namespace Plain.Core.Tests;

public static class AnnotationsSuite
{
    public static Suite Run()
    {
        var s = new Suite("annotations");
        if (Fixtures.Missing(s, "review.docx", "doc.docx")) return s;

        var file = PlainFile.Open(Fixtures.Path_("review.docx")!);

        var notes = Annotations.Comments(file);
        s.Equal("both comments are read", 2, notes.Count);
        s.Check("a comment carries its author", notes.Any(n => n.Author.Contains("Alex")));
        s.Check("and the other one's", notes.Any(n => n.Author.Contains("Sam")));
        s.Check("a comment carries its text", notes.Any(n => n.Text.Contains("thirty days right")));
        s.Check("a comment carries a readable date", notes.All(n => n.When.Length > 0));

        var revisions = Annotations.Revisions(file);
        s.Equal("both tracked changes are read", 2, revisions.Count);
        s.Check("one is an insertion", revisions.Any(r => r.Inserted));
        s.Check("one is a deletion", revisions.Any(r => !r.Inserted));
        s.Check("the inserted words are readable", revisions.Any(r => r.Inserted && r.Text.Contains("thirty")));
        s.Check("the deleted words are readable", revisions.Any(r => !r.Inserted && r.Text.Contains("sixty")));
        s.Check("a change names who made it", revisions.All(r => r.Author.Length > 0));

        // A document with none of either must say so quietly, not fall over.
        var plain = PlainFile.Open(Fixtures.Path_("doc.docx")!);
        s.Equal("a clean document has no comments", 0, Annotations.Comments(plain).Count);
        s.Equal("and no tracked changes", 0, Annotations.Revisions(plain).Count);

        // Settling the changes: what was added stays, what was struck out goes.
        var work = Fixtures.Copy("review.docx");
        try
        {
            var f = PlainFile.Open(work);
            string before = f.Document!.PlainText();
            s.Check("before, the inserted words are in the text", before.Contains("thirty (30) days"));

            int settled = Annotations.AcceptRevisions(f);
            s.Equal("both changes were settled", 2, settled);
            f.Save();

            var after = PlainFile.Open(work);
            s.Equal("no tracked changes are left", 0, Annotations.Revisions(after).Count);
            string text = after.Document!.PlainText();
            s.Check("the inserted words stayed", text.Contains("thirty (30) days"));
            s.Check("the struck out words went", !text.Contains("sixty (60) days"));
            s.Check("the rest of the document survived", text.Contains("invoice date"));
            s.Check("the document still opens as a document", after.Kind == FileKind.Document);
        }
        finally { try { File.Delete(work); } catch { } }

        // Clearing the comments.
        var second = Fixtures.Copy("review.docx");
        try
        {
            var f = PlainFile.Open(second);
            string before = f.Document!.PlainText();
            int removed = Annotations.RemoveComments(f);
            s.Equal("both comments were removed", 2, removed);
            f.Save();

            var after = PlainFile.Open(second);
            s.Equal("no comments are left", 0, Annotations.Comments(after).Count);
            s.Equal("the body text is untouched", before, after.Document!.PlainText());
            s.Check("the file still opens", after.Kind == FileKind.Document);
            s.Equal("removing them twice removes nothing the second time", 0, Annotations.RemoveComments(after));
        }
        finally { try { File.Delete(second); } catch { } }

        return s;
    }
}
