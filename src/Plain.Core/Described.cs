using System.Xml.Linq;

namespace Plain.Core;

/// <summary>
/// The words that describe a picture to somebody who cannot see it, and the words a presenter was going to say.
///
/// Both are text that lives in a file and is never on the page. Alt text is the one people are asked for and skip;
/// speaker notes are the one people forget is travelling with the deck they just emailed. Plain shows both and
/// lets both be changed, which is the same idea as the preserved rail: what is in there, said plainly.
/// </summary>
public static class Described
{
    /// <summary>A picture in the file, and what it is described as. An empty description is the point of this.</summary>
    public sealed record Described_(string Where, string Name, string Text);

    /// <summary>Every picture in the file, with its description, whether or not it has one.</summary>
    public static IReadOnlyList<Described_> Pictures(PlainFile file)
    {
        var found = new List<Described_>();
        foreach (var part in Parts(file))
        {
            if (!file.Package.Has(part.Part)) continue;
            var doc = Xml.Parse(file.Package.Read(part.Part));

            // Every shape carries a name and a description in the same element, whichever kind of file it is in.
            foreach (var properties in doc.Root!.Descendants().Where(x => x.Name.LocalName is "cNvPr" or "docPr"))
            {
                // Only the ones that are actually a picture; a text box has the same element.
                var owner = properties.Parent?.Parent;
                bool picture = owner?.Name.LocalName is "pic" or "inline" or "anchor"
                            || properties.Parent?.Name.LocalName == "nvPicPr"
                            || properties.Name.LocalName == "docPr";
                if (!picture) continue;

                var name = (string?)properties.Attribute("name") ?? "picture";
                var text = (string?)properties.Attribute("descr") ?? "";

                // One picture carries its name and description on two elements: the drawing's and the picture's
                // own. To a person that is one picture, so it is reported once, and described on both.
                var already = found.FindIndex(x => x.Where == part.Where && x.Name == name);
                if (already >= 0)
                {
                    if (found[already].Text.Length == 0 && text.Length > 0) found[already] = found[already] with { Text = text };
                    continue;
                }
                found.Add(new Described_(part.Where, name, text));
            }
        }
        return found;
    }

    /// <summary>
    /// Describe one picture. Both the elements that carry a description are set, because Word and PowerPoint each
    /// read a different one and a picture described in only one of them is described to only half the readers.
    /// Returns how many pictures were described, not how many elements were touched.
    /// </summary>
    public static int DescribePictures(PlainFile file, string where, string name, string text)
    {
        int changed = 0;
        foreach (var part in Parts(file))
        {
            if (part.Where != where || !file.Package.Has(part.Part)) continue;
            var doc = Xml.Parse(file.Package.Read(part.Part));

            bool touched = false;
            foreach (var properties in doc.Root!.Descendants().Where(x => x.Name.LocalName is "cNvPr" or "docPr"))
            {
                if (((string?)properties.Attribute("name") ?? "picture") != name) continue;
                properties.SetAttributeValue("descr", text.Length == 0 ? null : text);
                touched = true;
            }

            if (touched) { file.Package.Write(part.Part, Xml.ToBytes(doc)); changed++; }
        }
        return changed;
    }

    private sealed record Place(string Part, string Where);

    private static IEnumerable<Place> Parts(PlainFile file)
    {
        if (file.Document is not null && file.Package.Has(Document.BodyPart))
            yield return new Place(Document.BodyPart, "the text");
        foreach (var slide in file.Deck?.Slides ?? (IReadOnlyList<Slide>)Array.Empty<Slide>())
            yield return new Place(slide.PartName, $"slide {slide.Number}");
    }

    // ---------- what the presenter was going to say ----------

    /// <summary>One slide's notes: the slide it belongs to, the part they live in, and what they say.</summary>
    public sealed record Note(int Slide, string Part, string Text);

    /// <summary>The notes on every slide that has any.</summary>
    public static IReadOnlyList<Note> Notes(PlainFile file)
    {
        var found = new List<Note>();
        var deck = file.Deck;
        if (deck is null) return found;

        foreach (var slide in deck.Slides)
        {
            var part = NotesPartFor(file.Package, slide.PartName);
            if (part is null) continue;
            found.Add(new Note(slide.Number, part, TextIn(file.Package, part)));
        }
        return found;
    }

    /// <summary>Change what a slide's notes say. The first paragraph carries it and the rest are emptied.</summary>
    public static bool WriteNotes(OpcPackage pkg, string part, string text)
    {
        if (!pkg.Has(part)) return false;
        var doc = Xml.Parse(pkg.Read(part));
        var d = Dialect.Of(doc.Root!);
        XNamespace a = d.Draw;

        // The notes body is the text frame that is not the little picture of the slide.
        var paragraphs = doc.Root!.Descendants(a + "p").ToList();
        if (paragraphs.Count == 0) return false;

        var first = paragraphs[0];
        var keep = first.Element(a + "pPr");
        first.RemoveNodes();
        if (keep is not null) first.Add(keep);
        first.Add(new XElement(a + "r",
            new XElement(a + "rPr", new XAttribute("lang", "en-GB"), new XAttribute("dirty", "0")),
            new XElement(a + "t", text)));

        foreach (var other in paragraphs.Skip(1))
        {
            var theirs = other.Element(a + "pPr");
            other.RemoveNodes();
            if (theirs is not null) other.Add(theirs);
        }

        pkg.Write(part, Xml.ToBytes(doc));
        return true;
    }

    private static string? NotesPartFor(OpcPackage pkg, string slidePart)
    {
        var rels = new Rels(pkg, slidePart);
        for (int i = 1; i <= 40; i++)
        {
            var target = rels[$"rId{i}"];
            if (target is not null && target.Contains("notesSlide", StringComparison.Ordinal) && pkg.Has(target))
                return target;
        }
        return null;
    }

    private static string TextIn(OpcPackage pkg, string part)
    {
        var doc = Xml.Parse(pkg.Read(part));
        var d = Dialect.Of(doc.Root!);
        XNamespace a = d.Draw;
        var lines = doc.Root!.Descendants(a + "p")
            .Select(p => string.Concat(p.Descendants(a + "t").Select(t => t.Value)))
            .Where(x => x.Trim().Length > 0);
        return string.Join(Environment.NewLine, lines);
    }
}
