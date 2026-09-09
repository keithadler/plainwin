using System.Xml.Linq;

namespace Plain.Core;

/// <summary>One comment somebody left on a document, spreadsheet or deck.</summary>
public sealed record Note(string Author, string When, string Text, string Where, int Index = 0);

/// <summary>One tracked change: text somebody added or took out, and who.</summary>
public sealed record Revision(bool Inserted, string Author, string When, string Text, int Index = 0);

/// <summary>
/// What other people wrote in the margins. Plain preserves comments and tracked changes whether or not it shows
/// them, which for anyone sending a document onward is the wrong way round: they need to know what is in there
/// before it goes. So this reads them, and can clear them out.
/// </summary>
public static class Annotations
{
    // ---------- reading ----------

    public static IReadOnlyList<Note> Comments(PlainFile file) => file.Kind switch
    {
        FileKind.Document => WordComments(file.Package),
        FileKind.Spreadsheet => SheetComments(file.Package),
        FileKind.Presentation => DeckComments(file.Package),
        _ => Array.Empty<Note>(),
    };

    private static IReadOnlyList<Note> WordComments(OpcPackage pkg)
    {
        var notes = new List<Note>();
        if (!pkg.Has("word/comments.xml")) return notes;
        XDocument doc;
        try { doc = Xml.Parse(pkg.Read("word/comments.xml")); } catch { return notes; }
        var d = Dialect.Of(doc.Root!);

        foreach (var comment in doc.Root!.Elements(d.Word + "comment"))
        {
            var shape = new TextShape(comment, d.Word);
            string text = string.Join(" ", shape.Paragraphs.Select(shape.TextOf).Where(t => t.Length > 0));
            notes.Add(new Note(
                (string?)comment.Attribute(d.Word + "author") ?? "",
                When((string?)comment.Attribute(d.Word + "date")),
                text,
                "comment",
                notes.Count));
        }
        return notes;
    }

    private static IReadOnlyList<Note> SheetComments(OpcPackage pkg)
    {
        var notes = new List<Note>();

        // The old kind: a comment box anchored to a cell, with the authors in a list beside it.
        foreach (var part in pkg.Parts.Where(p => p.Name.StartsWith("xl/comments", StringComparison.Ordinal)).ToList())
        {
            XDocument doc;
            try { doc = Xml.Parse(pkg.Read(part.Name)); } catch { continue; }
            var d = Dialect.Of(doc.Root!);
            var authors = doc.Root!.Element(d.Sheet + "authors")?.Elements(d.Sheet + "author").Select(a => a.Value).ToList()
                          ?? new List<string>();
            foreach (var comment in doc.Root.Element(d.Sheet + "commentList")?.Elements(d.Sheet + "comment") ?? Enumerable.Empty<XElement>())
            {
                int who = Xml.Int(comment.Attribute("authorId"), -1);
                notes.Add(new Note(
                    who >= 0 && who < authors.Count ? authors[who] : "",
                    "",
                    string.Concat(comment.Descendants(d.Sheet + "t").Select(t => t.Value)).Trim(),
                    (string?)comment.Attribute("ref") ?? ""));
            }
        }

        // The newer kind, the ones that look like a conversation.
        foreach (var part in pkg.Parts.Where(p => p.Name.Contains("threadedComment", StringComparison.OrdinalIgnoreCase)).ToList())
        {
            XDocument doc;
            try { doc = Xml.Parse(pkg.Read(part.Name)); } catch { continue; }
            foreach (var comment in doc.Root!.Elements().Where(e => e.Name.LocalName == "threadedComment"))
                notes.Add(new Note(
                    "",
                    When((string?)comment.Attribute("dT")),
                    comment.Elements().FirstOrDefault(e => e.Name.LocalName == "text")?.Value ?? "",
                    (string?)comment.Attribute("ref") ?? ""));
        }
        return notes;
    }

    private static IReadOnlyList<Note> DeckComments(OpcPackage pkg)
    {
        var notes = new List<Note>();
        foreach (var part in pkg.Parts.Where(p => p.Name.Contains("/comments/", StringComparison.Ordinal)).ToList())
        {
            XDocument doc;
            try { doc = Xml.Parse(pkg.Read(part.Name)); } catch { continue; }
            foreach (var comment in doc.Root!.Elements().Where(e => e.Name.LocalName == "cm"))
                notes.Add(new Note(
                    "",
                    When((string?)comment.Attribute("dt")),
                    comment.Elements().FirstOrDefault(e => e.Name.LocalName == "text")?.Value ?? "",
                    part.Name));
        }
        return notes;
    }

    /// <summary>Every tracked change in a document, in the order they appear.</summary>
    public static IReadOnlyList<Revision> Revisions(PlainFile file)
    {
        var revisions = new List<Revision>();
        if (file.Kind != FileKind.Document || !file.Package.Has(Document.BodyPart)) return revisions;
        XDocument doc;
        try { doc = Xml.Parse(file.Package.Read(Document.BodyPart)); } catch { return revisions; }
        var d = Dialect.Of(doc.Root!);

        foreach (var change in doc.Descendants().Where(e => e.Name == d.Word + "ins" || e.Name == d.Word + "del"))
        {
            bool inserted = change.Name.LocalName == "ins";
            // Deleted text is kept in delText so it can be put back; inserted text is ordinary text.
            string text = string.Concat(change.Descendants()
                .Where(e => e.Name == d.Word + (inserted ? "t" : "delText"))
                .Select(e => e.Value));
            if (text.Length == 0) continue;
            revisions.Add(new Revision(inserted,
                (string?)change.Attribute(d.Word + "author") ?? "",
                When((string?)change.Attribute(d.Word + "date")),
                text,
                revisions.Count));
        }
        return revisions;
    }

    private static string When(string? raw)
    {
        if (string.IsNullOrWhiteSpace(raw)) return "";
        return DateTimeOffset.TryParse(raw, out var when) ? when.ToString("d MMM yyyy") : raw;
    }

    // ---------- clearing out ----------

    /// <summary>
    /// Settle every tracked change: keep what was added, drop what was struck out, and take the marks away. This is
    /// what "final" means, and it is what a document needs to be before it goes to somebody who should not see the
    /// argument that produced it. Returns how many changes were settled.
    /// </summary>
    public static int AcceptRevisions(PlainFile file) => Settle(file, accept: true, only: null);

    /// <summary>
    /// Turn down every tracked change: put back what was struck out, take out what was added. Reviewing mark-up is
    /// half saying no, and an editor that could only ever say yes was no use to anyone negotiating anything.
    /// </summary>
    public static int RejectRevisions(PlainFile file) => Settle(file, accept: false, only: null);

    /// <summary>Settle one change, by its place in the list <see cref="Revisions"/> gives back.</summary>
    public static bool SettleRevision(PlainFile file, int index, bool accept) => Settle(file, accept, index) > 0;

    private static int Settle(PlainFile file, bool accept, int? only)
    {
        if (file.Kind != FileKind.Document || !file.Package.Has(Document.BodyPart)) return 0;
        var doc = Xml.Parse(file.Package.Read(Document.BodyPart));
        var d = Dialect.Of(doc.Root!);
        int settled = 0, seen = 0;

        // Walk in document order, the same order Revisions reports, so an index means the same thing to both.
        foreach (var change in doc.Descendants().Where(e => e.Name == d.Word + "ins" || e.Name == d.Word + "del").ToList())
        {
            bool inserted = change.Name.LocalName == "ins";
            string text = string.Concat(change.Descendants()
                .Where(e => e.Name == d.Word + (inserted ? "t" : "delText")).Select(e => e.Value));
            if (text.Length == 0) continue;      // Revisions skips these, so the numbering must too

            int at = seen++;
            if (only is { } wanted && at != wanted) continue;

            bool keep = inserted == accept;
            if (!keep) { change.Remove(); }
            else
            {
                var kept = change.Elements().ToList();
                foreach (var run in kept) run.Remove();
                // Text that was struck out is stored as delText; putting it back makes it ordinary text again.
                if (!inserted)
                    foreach (var deleted in kept.SelectMany(r => r.Elements(d.Word + "delText").ToList()))
                        deleted.ReplaceWith(new XElement(d.Word + "t",
                            new XAttribute(XNamespace.Xml + "space", "preserve"), deleted.Value));
                change.ReplaceWith(kept);
            }
            settled++;
            if (only is not null) break;
        }

        if (only is null)
        {
            // A paragraph mark can itself be marked as changed; that mark goes too.
            foreach (var mark in doc.Descendants(d.Word + "rPr").Elements(d.Word + "ins").ToList()) mark.Remove();
            foreach (var mark in doc.Descendants(d.Word + "rPr").Elements(d.Word + "del").ToList()) mark.Remove();
        }

        if (settled > 0) file.Package.Write(Document.BodyPart, Xml.ToBytes(doc));
        return settled;
    }

    /// <summary>Take out one comment, by its place in the list <see cref="Comments"/> gives back.</summary>
    public static bool RemoveComment(PlainFile file, int index)
    {
        if (file.Kind != FileKind.Document || !file.Package.Has("word/comments.xml")) return false;
        var comments = Xml.Parse(file.Package.Read("word/comments.xml"));
        var cd = Dialect.Of(comments.Root!);
        var all = comments.Root!.Elements(cd.Word + "comment").ToList();
        if (index < 0 || index >= all.Count) return false;

        string? id = (string?)all[index].Attribute(cd.Word + "id");
        all[index].Remove();
        file.Package.Write("word/comments.xml", Xml.ToBytes(comments));

        if (id is not null)
        {
            // Take the marks that pointed at it out of the document, and leave the others alone.
            var body = Xml.Parse(file.Package.Read(Document.BodyPart));
            var d = Dialect.Of(body.Root!);
            foreach (var name in new[] { "commentRangeStart", "commentRangeEnd", "commentReference" })
                foreach (var mark in body.Descendants(d.Word + name).ToList())
                    if ((string?)mark.Attribute(d.Word + "id") == id) mark.Remove();
            foreach (var run in body.Descendants(d.Word + "r").ToList())
                if (!run.Elements().Any(e => e.Name != d.Word + "rPr")) run.Remove();
            file.Package.Write(Document.BodyPart, Xml.ToBytes(body));
        }
        return true;
    }

    /// <summary>
    /// Take every comment out: the text of them, and the marks in the document that point at them. Returns how many
    /// went. The comments part is emptied rather than removed, because Plain never takes a part out of a file.
    /// </summary>
    public static int RemoveComments(PlainFile file)
    {
        int removed = 0;

        if (file.Kind == FileKind.Document && file.Package.Has("word/comments.xml"))
        {
            var comments = Xml.Parse(file.Package.Read("word/comments.xml"));
            var cd = Dialect.Of(comments.Root!);
            removed = comments.Root!.Elements(cd.Word + "comment").Count();
            if (removed > 0)
            {
                comments.Root.Elements(cd.Word + "comment").Remove();
                file.Package.Write("word/comments.xml", Xml.ToBytes(comments));

                var body = Xml.Parse(file.Package.Read(Document.BodyPart));
                var d = Dialect.Of(body.Root!);
                foreach (var name in new[] { "commentRangeStart", "commentRangeEnd", "commentReference" })
                    body.Descendants(d.Word + name).ToList().ForEach(e => e.Remove());
                // A run that held nothing but a comment mark is now an empty run; take it out too.
                foreach (var run in body.Descendants(d.Word + "r").ToList())
                    if (!run.Elements().Any(e => e.Name != d.Word + "rPr")) run.Remove();
                file.Package.Write(Document.BodyPart, Xml.ToBytes(body));
            }
        }

        foreach (var part in file.Package.Parts
                     .Where(p => p.Name.Contains("threadedComment", StringComparison.OrdinalIgnoreCase)
                              || p.Name.StartsWith("xl/comments", StringComparison.Ordinal)
                              || p.Name.Contains("/comments/", StringComparison.Ordinal)).ToList())
        {
            XDocument doc;
            try { doc = Xml.Parse(file.Package.Read(part.Name)); } catch { continue; }
            var children = doc.Root!.Elements()
                .Where(e => e.Name.LocalName is "threadedComment" or "cm" or "commentList").ToList();
            int here = children.Sum(c => c.Name.LocalName == "commentList" ? c.Elements().Count() : 1);
            if (here == 0) continue;
            foreach (var child in children)
                if (child.Name.LocalName == "commentList") child.Elements().Remove(); else child.Remove();
            file.Package.Write(part.Name, Xml.ToBytes(doc));
            removed += here;
        }

        return removed;
    }
}
