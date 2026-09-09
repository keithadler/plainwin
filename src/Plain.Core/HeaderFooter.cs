using System.Xml.Linq;

namespace Plain.Core;

/// <summary>
/// The lines that run along the top and bottom of every page of a document.
///
/// They are the part of a file people forget is there. A header saying which client a template was written for
/// travels with every copy of it, and nobody sees it on screen because Word only shows it on the page. Plain shows
/// what they say and lets the words be changed.
///
/// A document can have several: a different one for the first page, for left and right pages, and one for the
/// rest. Each is a part of its own. Plain lists them all rather than showing one and hiding the others, because
/// the one it hid would be the one that mattered.
/// </summary>
public static class HeaderFooter
{
    /// <summary>One header or footer: which it is, which pages it is for, and what it says.</summary>
    public sealed record Band(string Part, bool IsHeader, string Which, string Text);

    private const string HeaderType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/header";
    private const string FooterType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer";

    /// <summary>Every header and footer the document has, in the order Word lists them.</summary>
    public static IReadOnlyList<Band> All(PlainFile file)
    {
        var found = new List<Band>();
        var pkg = file.Package;
        if (file.Document is null || !pkg.Has(Document.BodyPart)) return found;

        var doc = Xml.Parse(pkg.Read(Document.BodyPart));
        var d = Dialect.Of(doc.Root!);
        var rels = new Rels(pkg, Document.BodyPart);

        var body = doc.Root!.Element(d.Word + "body");
        foreach (var section in body?.Descendants(d.Word + "sectPr") ?? Enumerable.Empty<XElement>())
            foreach (var reference in section.Elements())
            {
                bool header = reference.Name == d.Word + "headerReference";
                bool footer = reference.Name == d.Word + "footerReference";
                if (!header && !footer) continue;

                var id = (string?)reference.Attribute(d.Rel + "id");
                var part = id is null ? null : rels[id];
                if (part is null || !pkg.Has(part)) continue;
                if (found.Any(b => b.Part == part)) continue;

                found.Add(new Band(part, header, Which((string?)reference.Attribute(d.Word + "type")), Read(pkg, part)));
            }

        return found;
    }

    private static string Which(string? type) => type switch
    {
        "first" => "the first page",
        "even" => "left-hand pages",
        _ => "every page",
    };

    /// <summary>What one says, as plain text, with each paragraph on its own line.</summary>
    public static string Read(OpcPackage pkg, string part)
    {
        if (!pkg.Has(part)) return "";
        var doc = Xml.Parse(pkg.Read(part));
        var d = Dialect.Of(doc.Root!);
        var lines = doc.Root!.Descendants(d.Word + "p")
            .Select(p => string.Concat(p.Descendants(d.Word + "t").Select(t => t.Value)));
        return string.Join(Environment.NewLine, lines).TrimEnd();
    }

    /// <summary>
    /// Change what one says. The first paragraph carries the new text and any others are emptied, because a header
    /// people can edit here has to end up saying exactly what they typed and nothing left over underneath it.
    /// Everything about how it looks is left alone: this replaces words, not formatting.
    /// </summary>
    public static bool Write(OpcPackage pkg, string part, string text)
    {
        if (!pkg.Has(part)) return false;
        var doc = Xml.Parse(pkg.Read(part));
        var d = Dialect.Of(doc.Root!);

        var paragraphs = doc.Root!.Descendants(d.Word + "p").ToList();
        if (paragraphs.Count == 0) return false;

        var first = paragraphs[0];
        // Keep the paragraph's own properties, which say how it is laid out, and replace only what it says.
        var properties = first.Element(d.Word + "pPr");
        var runProperties = first.Elements(d.Word + "r").FirstOrDefault()?.Element(d.Word + "rPr");
        first.RemoveNodes();
        if (properties is not null) first.Add(properties);

        var run = new XElement(d.Word + "r");
        if (runProperties is not null) run.Add(new XElement(runProperties));
        // xml:space is what stops Word trimming the spaces someone meant to type.
        run.Add(new XElement(d.Word + "t", new XAttribute(XNamespace.Xml + "space", "preserve"), text));
        first.Add(run);

        foreach (var other in paragraphs.Skip(1))
        {
            var keep = other.Element(d.Word + "pPr");
            other.RemoveNodes();
            if (keep is not null) other.Add(keep);
        }

        pkg.Write(part, Xml.ToBytes(doc));
        return true;
    }
}
