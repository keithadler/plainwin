using System.Xml.Linq;

namespace Plain.Core;

/// <summary>
/// Changing the shape of a deck: moving a slide, taking one out, and putting a new empty one in.
///
/// A deck keeps its running order in one list in presentation.xml, and each entry points at a slide part through a
/// relationship. Moving a slide is only reordering that list, which is why it is completely safe. Removing one
/// takes the entry and the relationship out and leaves the part itself in the package: an unreferenced part costs
/// a few kilobytes and nothing else, whereas deleting a part that something unnoticed still points at is how a
/// deck gets broken. Plain would rather leave a harmless orphan than risk that.
///
/// Adding a slide writes a new part modelled on the empty deck Plain already knows how to build, borrowing the
/// layout the deck's first slide uses so the new one inherits the same look.
/// </summary>
public static class Slides
{
    private const string DocRel = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    private const string PresPart = "ppt/presentation.xml";

    public abstract record Result;
    public sealed record Done(string What) : Result;
    public sealed record Refused(string Reason) : Result;

    /// <summary>Move the slide at <paramref name="from"/> (counting from 1) to <paramref name="to"/>.</summary>
    public static Result Move(OpcPackage pkg, int from, int to)
    {
        var doc = Xml.Parse(pkg.Read(PresPart));
        var d = Dialect.Of(doc.Root!);
        var list = doc.Root?.Element(d.Pres + "sldIdLst");
        var ids = list?.Elements(d.Pres + "sldId").ToList();
        if (list is null || ids is null || ids.Count == 0) return new Refused("This deck has no slides to move.");

        if (from < 1 || from > ids.Count) return new Refused($"There is no slide {from}.");
        to = Math.Clamp(to, 1, ids.Count);
        if (from == to) return new Done("That slide is already there.");

        var moving = ids[from - 1];
        moving.Remove();
        var rest = list.Elements(d.Pres + "sldId").ToList();
        if (to - 1 >= rest.Count) rest[^1].AddAfterSelf(moving);
        else rest[to - 1].AddBeforeSelf(moving);

        pkg.Write(PresPart, Xml.ToBytes(doc));
        return new Done($"Slide {from} is now slide {to}.");
    }

    /// <summary>
    /// Take a slide out of the running order. The part stays in the package unreferenced, which is harmless, and
    /// means nothing that still points at it can break.
    /// </summary>
    public static Result Remove(OpcPackage pkg, int number)
    {
        var doc = Xml.Parse(pkg.Read(PresPart));
        var d = Dialect.Of(doc.Root!);
        var list = doc.Root?.Element(d.Pres + "sldIdLst");
        var ids = list?.Elements(d.Pres + "sldId").ToList();
        if (list is null || ids is null) return new Refused("This deck has no slides.");
        if (ids.Count <= 1) return new Refused("A deck has to have a slide in it, so this one cannot be removed.");
        if (number < 1 || number > ids.Count) return new Refused($"There is no slide {number}.");

        ids[number - 1].Remove();
        pkg.Write(PresPart, Xml.ToBytes(doc));
        return new Done($"Slide {number} is out of the deck. Its contents stay in the file, unused.");
    }

    /// <summary>Put a new empty slide after <paramref name="after"/>, or at the front when that is nought.</summary>
    public static Result Add(OpcPackage pkg, int after)
    {
        var doc = Xml.Parse(pkg.Read(PresPart));
        var d = Dialect.Of(doc.Root!);
        var list = doc.Root?.Element(d.Pres + "sldIdLst");
        if (list is null) return new Refused("This deck has no list of slides to add to.");

        var rels = new Rels(pkg, PresPart);
        var layout = FirstLayout(pkg);
        if (layout is null) return new Refused("Plain could not find a slide layout to base a new slide on.");

        // A part name and a relationship id nothing else is using.
        int n = 1;
        while (pkg.Has($"ppt/slides/slide{n}.xml")) n++;
        string part = $"ppt/slides/slide{n}.xml";

        int rid = 1;
        while (rels[$"rId{rid}"] is not null) rid++;
        string id = $"rId{rid}";

        string ns = "http://schemas.openxmlformats.org/drawingml/2006/main";
        string p = "http://schemas.openxmlformats.org/presentationml/2006/main";
        string slide =
            "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
            $"<p:sld xmlns:a=\"{ns}\" xmlns:r=\"{DocRel}\" xmlns:p=\"{p}\"><p:cSld><p:spTree>" +
            "<p:nvGrpSpPr><p:cNvPr id=\"1\" name=\"\"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>" +
            "<p:grpSpPr/></p:spTree></p:cSld><p:clrMapOvr><a:overrideClrMapping bg1=\"lt1\" tx1=\"dk1\" " +
            "bg2=\"lt2\" tx2=\"dk2\" accent1=\"accent1\" accent2=\"accent2\" accent3=\"accent3\" accent4=\"accent4\" " +
            "accent5=\"accent5\" accent6=\"accent6\" hlink=\"hlink\" folHlink=\"folHlink\"/></p:clrMapOvr></p:sld>";
        pkg.Add(part, System.Text.Encoding.UTF8.GetBytes(slide));

        // Its own relationship to a layout, or PowerPoint will not open it.
        string relative = layout.StartsWith("ppt/", StringComparison.Ordinal) ? "../" + layout["ppt/".Length..] : layout;
        pkg.Add($"ppt/slides/_rels/slide{n}.xml.rels", System.Text.Encoding.UTF8.GetBytes(
            "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
            "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">" +
            $"<Relationship Id=\"rId1\" Type=\"{DocRel}/slideLayout\" Target=\"{relative}\"/></Relationships>"));

        AddRelationship(pkg, PresPart, id, $"{DocRel}/slide", $"slides/slide{n}.xml");
        AddContentType(pkg, "/" + part,
            "application/vnd.openxmlformats-officedocument.presentationml.slide+xml");

        // A slide id has to be at least 256 and unique.
        int newId = 256;
        foreach (var existing in list.Elements(d.Pres + "sldId"))
            newId = Math.Max(newId, Xml.Int(existing.Attribute("id"), 256) + 1);

        var entry = new XElement(d.Pres + "sldId",
            new XAttribute("id", newId),
            new XAttribute(d.Rel + "id", id));

        var all = list.Elements(d.Pres + "sldId").ToList();
        if (after <= 0 || all.Count == 0) list.AddFirst(entry);
        else if (after >= all.Count) all[^1].AddAfterSelf(entry);
        else all[after - 1].AddAfterSelf(entry);

        pkg.Write(PresPart, Xml.ToBytes(doc));
        return new Done($"A new empty slide is in at {Math.Clamp(after + 1, 1, all.Count + 1)}.");
    }

    /// <summary>
    /// The layout an existing slide uses, so a new slide inherits the look of the ones already there. Falls back to
    /// any layout in the package, because a slide with no layout is one PowerPoint will not open.
    /// </summary>
    private static string? FirstLayout(OpcPackage pkg)
    {
        foreach (var part in pkg.Parts.Select(x => x.Name)
                     .Where(x => x.StartsWith("ppt/slides/slide", StringComparison.Ordinal)
                              && x.EndsWith(".xml", StringComparison.Ordinal)))
        {
            var rels = new Rels(pkg, part);
            for (int i = 1; i <= 30; i++)
            {
                var target = rels[$"rId{i}"];
                if (target is not null && target.StartsWith("ppt/slideLayouts/", StringComparison.Ordinal)
                    && pkg.Has(target)) return target;
            }
        }
        return pkg.Parts.Select(x => x.Name)
                  .FirstOrDefault(x => x.StartsWith("ppt/slideLayouts/", StringComparison.Ordinal)
                                    && x.EndsWith(".xml", StringComparison.Ordinal));
    }

    private static void AddRelationship(OpcPackage pkg, string partName, string id, string type, string target)
    {
        // The _rels part that belongs to this one: alongside it, in a _rels folder, with .rels on the end.
        int slash = partName.LastIndexOf('/');
        string relsPart = (slash < 0 ? "" : partName[..(slash + 1)]) + "_rels/" + partName[(slash + 1)..] + ".rels";
        var doc = pkg.Has(relsPart)
            ? Xml.Parse(pkg.Read(relsPart))
            : XDocument.Parse("<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"/>");
        XNamespace r = "http://schemas.openxmlformats.org/package/2006/relationships";
        doc.Root!.Add(new XElement(r + "Relationship",
            new XAttribute("Id", id), new XAttribute("Type", type), new XAttribute("Target", target)));
        pkg.Write(relsPart, Xml.ToBytes(doc));
    }

    private static void AddContentType(OpcPackage pkg, string partName, string type)
    {
        var doc = Xml.Parse(pkg.Read("[Content_Types].xml"));
        XNamespace ct = "http://schemas.openxmlformats.org/package/2006/content-types";
        if (doc.Root!.Elements(ct + "Override").Any(o => (string?)o.Attribute("PartName") == partName)) return;
        doc.Root.Add(new XElement(ct + "Override",
            new XAttribute("PartName", partName), new XAttribute("ContentType", type)));
        pkg.Write("[Content_Types].xml", Xml.ToBytes(doc));
    }
}
