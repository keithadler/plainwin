using System.Xml.Linq;

namespace Plain.Core;

/// <summary>
/// Putting something into a file that was not there: a link on some words, and a picture.
///
/// These are the two places Plain adds rather than edits, and both are done the narrow way. A link is put on a
/// whole paragraph rather than on a chosen run of letters, because choosing letters means modelling how a
/// paragraph is broken into runs, and getting that wrong damages the text. A picture goes at the end of the
/// document or on a slide at a stated size, rather than being placed by dragging.
///
/// Both say what they did. Neither pretends to be Word.
///
/// Both write into the package rather than through the model that is open on it, so both put the model's own
/// changes into the package first and expect the caller to open the file again afterwards. Carrying on with the
/// model that was open would write its older copy of the text back over what these just did.
/// </summary>
public static class Insert
{
    public abstract record Result;
    public sealed record Done(string What) : Result;
    public sealed record Refused(string Reason) : Result;

    private const string DocRel = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    private const string LinkType = DocRel + "/hyperlink";
    private const string ImageType = DocRel + "/image";

    /// <summary>
    /// Make a paragraph of a document into a link. Only ordinary web and mail addresses: Plain will not write a
    /// link into a file that it would refuse to open itself.
    /// </summary>
    public static Result Link(PlainFile file, int block, string address)
    {
        if (file.Document is null) return new Refused("Links can be added to a document.");
        // The model holds its own copy of the body; putting its edits into the package first means this does not
        // undo them, and the caller reopening afterwards means the model does not undo this.
        file.Flush();
        if (!Links.SafeToOpen(address))
            return new Refused("Plain only writes ordinary web and mail addresses, the ones it would open itself. "
                             + "Anything that could run a program is not something to put in a document.");

        var pkg = file.Package;
        var doc = Xml.Parse(pkg.Read(Document.BodyPart));
        var d = Dialect.Of(doc.Root!);
        var body = doc.Root!.Element(d.Word + "body") ?? doc.Root!;

        var paragraphs = body.Descendants(d.Word + "p").ToList();
        if (block < 0 || block >= paragraphs.Count) return new Refused($"There is no paragraph {block + 1}.");
        var paragraph = paragraphs[block];

        var runs = paragraph.Elements(d.Word + "r").ToList();
        if (runs.Count == 0) return new Refused("That paragraph has no words to put a link on.");
        if (paragraph.Elements().Any(x => x.Name == d.Word + "hyperlink"))
            return new Refused("That paragraph already has a link on it.");

        string id = AddRelationship(pkg, Document.BodyPart, LinkType, address, external: true);

        // The runs move inside the link, keeping how they look; only who owns them changes.
        var link = new XElement(d.Word + "hyperlink", new XAttribute(d.Rel + "id", id));
        runs[0].AddBeforeSelf(link);
        foreach (var run in runs) { run.Remove(); link.Add(run); }

        // A link nobody can see is a link nobody knows is there, so the words take the document's own Hyperlink
        // style where it has one. Where it has none, Plain says so rather than inventing a blue underline: making
        // up a style would be Plain deciding how the document should look, which is not its job.
        bool styled = false;
        if (HasCharacterStyle(pkg, "Hyperlink"))
            foreach (var run in link.Elements(d.Word + "r"))
            {
                var properties = run.Element(d.Word + "rPr");
                if (properties is null) { properties = new XElement(d.Word + "rPr"); run.AddFirst(properties); }
                properties.Elements(d.Word + "rStyle").Remove();
                properties.AddFirst(new XElement(d.Word + "rStyle", new XAttribute(d.Word + "val", "Hyperlink")));
                styled = true;
            }

        pkg.Write(Document.BodyPart, Xml.ToBytes(doc));
        return new Done($"That paragraph now links to {address}."
                      + (styled ? "" : " This document has no link styling, so the words look as they did."));
    }

    /// <summary>Take the link off a paragraph, leaving the words where they are.</summary>
    public static Result Unlink(PlainFile file, int block)
    {
        if (file.Document is null) return new Refused("Links live in a document.");
        file.Flush();
        var pkg = file.Package;
        var doc = Xml.Parse(pkg.Read(Document.BodyPart));
        var d = Dialect.Of(doc.Root!);
        var body = doc.Root!.Element(d.Word + "body") ?? doc.Root!;

        var paragraphs = body.Descendants(d.Word + "p").ToList();
        if (block < 0 || block >= paragraphs.Count) return new Refused($"There is no paragraph {block + 1}.");

        var links = paragraphs[block].Elements(d.Word + "hyperlink").ToList();
        if (links.Count == 0) return new Refused("There is no link on that paragraph.");

        foreach (var link in links)
        {
            foreach (var run in link.Elements().ToList()) { run.Remove(); link.AddBeforeSelf(run); }
            link.Remove();
        }

        // The relationship is left in place. An unused one is harmless; removing one something else still uses
        // is not, and Plain cannot be sure nothing else does.
        pkg.Write(Document.BodyPart, Xml.ToBytes(doc));
        return new Done("The words are still there; the link is off them.");
    }

    /// <summary>The picture kinds Plain will put in, and what a package calls them.</summary>
    private static readonly Dictionary<string, string> Kinds = new(StringComparer.OrdinalIgnoreCase)
    {
        [".png"] = "image/png", [".jpg"] = "image/jpeg", [".jpeg"] = "image/jpeg",
        [".gif"] = "image/gif", [".bmp"] = "image/bmp",
    };

    /// <summary>
    /// Put a picture at the end of a document, at a size in centimetres. Nothing is scaled or re-encoded: the
    /// bytes on disk are the bytes that go in, so a picture Plain put in is the picture you gave it.
    /// </summary>
    public static Result Picture(PlainFile file, string path, double widthCm)
    {
        if (file.Document is null) return new Refused("Plain puts pictures into a document.");
        if (!File.Exists(path)) return new Refused($"There is no file at {path}.");
        file.Flush();

        var extension = System.IO.Path.GetExtension(path);
        if (!Kinds.TryGetValue(extension, out var kind))
            return new Refused($"Plain puts in {string.Join(", ", Kinds.Keys)} pictures. That one is {extension}.");

        var bytes = File.ReadAllBytes(path);
        if (bytes.LongLength > 32 * 1024 * 1024) return new Refused("That picture is larger than 32 MB.");

        var pkg = file.Package;
        int n = 1;
        while (pkg.Has($"word/media/plain{n}{extension}")) n++;
        string part = $"word/media/plain{n}{extension}";

        pkg.Add(part, bytes);
        AddContentType(pkg, extension.TrimStart('.'), kind);
        string id = AddRelationship(pkg, Document.BodyPart, ImageType, $"media/plain{n}{extension}", external: false);

        var doc = Xml.Parse(pkg.Read(Document.BodyPart));
        var d = Dialect.Of(doc.Root!);
        var body = doc.Root!.Element(d.Word + "body") ?? doc.Root!;

        // A twentieth of a point per unit, which is what the format counts in.
        long width = (long)Math.Round(Math.Clamp(widthCm, 0.5, 40) * 360000);
        long height = width * 3 / 4;

        XNamespace wp = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing";
        XNamespace a = d.Draw;
        XNamespace pic = "http://schemas.openxmlformats.org/drawingml/2006/picture";

        var drawing = new XElement(d.Word + "drawing",
            new XElement(wp + "inline", new XAttribute("distT", 0), new XAttribute("distB", 0),
                new XAttribute("distL", 0), new XAttribute("distR", 0),
                new XElement(wp + "extent", new XAttribute("cx", width), new XAttribute("cy", height)),
                new XElement(wp + "docPr", new XAttribute("id", 1000 + n),
                    new XAttribute("name", System.IO.Path.GetFileName(path))),
                new XElement(a + "graphic",
                    new XElement(a + "graphicData",
                        new XAttribute("uri", "http://schemas.openxmlformats.org/drawingml/2006/picture"),
                        new XElement(pic + "pic",
                            new XElement(pic + "nvPicPr",
                                new XElement(pic + "cNvPr", new XAttribute("id", 0),
                                    new XAttribute("name", System.IO.Path.GetFileName(path))),
                                new XElement(pic + "cNvPicPr")),
                            new XElement(pic + "blipFill",
                                new XElement(a + "blip", new XAttribute(d.Rel + "embed", id)),
                                new XElement(a + "stretch", new XElement(a + "fillRect"))),
                            new XElement(pic + "spPr",
                                new XElement(a + "xfrm",
                                    new XElement(a + "off", new XAttribute("x", 0), new XAttribute("y", 0)),
                                    new XElement(a + "ext", new XAttribute("cx", width), new XAttribute("cy", height))),
                                new XElement(a + "prstGeom", new XAttribute("prst", "rect"),
                                    new XElement(a + "avLst"))))))));

        var paragraph = new XElement(d.Word + "p", new XElement(d.Word + "r", drawing));
        // Before the section properties, which have to stay last in the body.
        var section = body.Elements(d.Word + "sectPr").FirstOrDefault();
        if (section is not null) section.AddBeforeSelf(paragraph); else body.Add(paragraph);

        pkg.Write(Document.BodyPart, Xml.ToBytes(doc));
        return new Done($"{System.IO.Path.GetFileName(path)} is at the end of the document, {widthCm:0.#} cm across.");
    }

    /// <summary>Does the document define a character style by this name? Applying one it lacks would do nothing.</summary>
    private static bool HasCharacterStyle(OpcPackage pkg, string id)
    {
        if (!pkg.Has("word/styles.xml")) return false;
        try
        {
            var styles = Xml.Parse(pkg.Read("word/styles.xml"));
            var d = Dialect.Of(styles.Root!);
            return styles.Root!.Elements(d.Word + "style")
                .Any(x => string.Equals((string?)x.Attribute(d.Word + "styleId"), id, StringComparison.OrdinalIgnoreCase));
        }
        catch { return false; }
    }

    private static string AddRelationship(OpcPackage pkg, string partName, string type, string target, bool external)
    {
        int slash = partName.LastIndexOf('/');
        string relsPart = partName[..(slash + 1)] + "_rels/" + partName[(slash + 1)..] + ".rels";
        XNamespace r = "http://schemas.openxmlformats.org/package/2006/relationships";

        XDocument doc;
        bool had = pkg.Has(relsPart);
        doc = had ? Xml.Parse(pkg.Read(relsPart))
                  : new XDocument(new XElement(r + "Relationships"));

        int n = 1;
        var taken = doc.Root!.Elements(r + "Relationship").Select(x => (string?)x.Attribute("Id")).ToHashSet();
        while (taken.Contains($"rId{n}")) n++;
        string id = $"rId{n}";

        var element = new XElement(r + "Relationship",
            new XAttribute("Id", id), new XAttribute("Type", type), new XAttribute("Target", target));
        if (external) element.SetAttributeValue("TargetMode", "External");
        doc.Root.Add(element);

        if (had) pkg.Write(relsPart, Xml.ToBytes(doc)); else pkg.Add(relsPart, Xml.ToBytes(doc));
        return id;
    }

    private static void AddContentType(OpcPackage pkg, string extension, string type)
    {
        var doc = Xml.Parse(pkg.Read("[Content_Types].xml"));
        XNamespace ct = "http://schemas.openxmlformats.org/package/2006/content-types";
        if (doc.Root!.Elements(ct + "Default").Any(x =>
                string.Equals((string?)x.Attribute("Extension"), extension, StringComparison.OrdinalIgnoreCase)))
            return;
        doc.Root.AddFirst(new XElement(ct + "Default",
            new XAttribute("Extension", extension), new XAttribute("ContentType", type)));
        pkg.Write("[Content_Types].xml", Xml.ToBytes(doc));
    }
}
