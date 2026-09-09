using System.Xml.Linq;

namespace Plain.Core;

/// <summary>
/// The links in a file, and where they actually go.
///
/// A link's words and its destination are two different things, and a file is one of the easier places to hide
/// that they disagree. "our terms" can point anywhere. Plain shows both, side by side, so what a document is
/// really offering can be read at a glance rather than hovered over one link at a time.
///
/// A link's address lives in the relationships part rather than in the text, which is why this reads both.
/// </summary>
public static class Links
{
    /// <summary>One link: the words it wears, where it goes, and where in the file it is.</summary>
    public sealed record Link(string Text, string Target, string Where, bool External);

    private const string LinkType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink";

    /// <summary>Every link in the file, whichever of the three kinds it is.</summary>
    public static IReadOnlyList<Link> All(PlainFile file) => file.Kind switch
    {
        FileKind.Document => InPart(file, Document.BodyPart, "the text"),
        FileKind.Presentation => InSlides(file),
        FileKind.Spreadsheet => InSheets(file),
        _ => Array.Empty<Link>(),
    };

    private static IReadOnlyList<Link> InPart(PlainFile file, string part, string where)
    {
        var found = new List<Link>();
        var pkg = file.Package;
        if (!pkg.Has(part)) return found;

        var doc = Xml.Parse(pkg.Read(part));
        var d = Dialect.Of(doc.Root!);
        var targets = External(pkg, part);

        foreach (var link in doc.Root!.Descendants().Where(x => x.Name.LocalName == "hyperlink"))
        {
            var id = (string?)link.Attribute(d.Rel + "id");
            var text = string.Concat(link.Descendants().Where(x => x.Name.LocalName == "t").Select(x => x.Value));
            // A link with no id points inside the document, at a bookmark, which is not a link out of it.
            var to = id is not null && targets.TryGetValue(id, out var address) ? address
                   : (string?)link.Attribute(d.Word + "anchor") is { } anchor ? "in this file: " + anchor
                   : "";
            if (text.Length == 0 && to.Length == 0) continue;
            found.Add(new Link(text, to, where, id is not null && targets.ContainsKey(id)));
        }
        return found;
    }

    private static IReadOnlyList<Link> InSlides(PlainFile file)
    {
        var found = new List<Link>();
        foreach (var slide in file.Deck?.Slides ?? (IReadOnlyList<Slide>)Array.Empty<Slide>())
            found.AddRange(InPart(file, slide.PartName, $"slide {slide.Number}"));
        return found;
    }

    /// <summary>
    /// A workbook keeps its links in the sheet rather than around the words, so the cell is what identifies one.
    /// </summary>
    private static IReadOnlyList<Link> InSheets(PlainFile file)
    {
        var found = new List<Link>();
        var pkg = file.Package;
        foreach (var sheet in file.Workbook?.Sheets ?? (IReadOnlyList<Sheet>)Array.Empty<Sheet>())
        {
            if (!pkg.Has(sheet.PartName)) continue;
            var doc = Xml.Parse(pkg.Read(sheet.PartName));
            var d = Dialect.Of(doc.Root!);
            var targets = External(pkg, sheet.PartName);

            foreach (var link in doc.Root!.Descendants(d.Sheet + "hyperlink"))
            {
                var cell = (string?)link.Attribute("ref") ?? "";
                var id = (string?)link.Attribute(d.Rel + "id");
                var to = id is not null && targets.TryGetValue(id, out var address) ? address
                       : (string?)link.Attribute("location") is { } place ? "in this file: " + place
                       : "";
                var text = CellRef.TryParse(cell, out var at) ? sheet.Read(at).Display : cell;
                found.Add(new Link(text, to, $"{sheet.Name}!{cell}", id is not null && targets.ContainsKey(id)));
            }
        }
        return found;
    }

    /// <summary>The addresses in a part's relationships that are links out of the file.</summary>
    private static Dictionary<string, string> External(OpcPackage pkg, string part)
    {
        var map = new Dictionary<string, string>(StringComparer.Ordinal);
        int slash = part.LastIndexOf('/');
        string relsPart = part[..(slash + 1)] + "_rels/" + part[(slash + 1)..] + ".rels";
        if (!pkg.Has(relsPart)) return map;

        XNamespace pkgNs = "http://schemas.openxmlformats.org/package/2006/relationships";
        foreach (var r in Xml.Parse(pkg.Read(relsPart)).Root!.Elements(pkgNs + "Relationship"))
        {
            if ((string?)r.Attribute("Type") != LinkType) continue;
            var id = (string?)r.Attribute("Id");
            var target = (string?)r.Attribute("Target");
            if (id is not null && target is not null) map[id] = target;
        }
        return map;
    }

    /// <summary>
    /// Is this address one Plain will offer to open? Only the ordinary web schemes and mail. Anything else, and
    /// in particular anything that would run a program, is shown but never opened, because a document is not a
    /// thing you should be able to click and have something happen.
    /// </summary>
    public static bool SafeToOpen(string target)
    {
        if (!Uri.TryCreate(target, UriKind.Absolute, out var uri)) return false;
        return uri.Scheme is "http" or "https" or "mailto";
    }
}
