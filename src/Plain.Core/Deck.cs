using System.Xml.Linq;

namespace Plain.Core;

/// <summary>One text frame on a slide: the placeholder it fills and the lines it holds.</summary>
public sealed record SlideText(int ShapeIndex, string Placeholder, IReadOnlyList<string> Lines);

/// <summary>
/// A PowerPoint deck, opened through <see cref="OpcPackage"/>. Plain reads and edits the text in each slide's shapes.
/// Pictures, diagrams, layouts, transitions and animations are never modelled, so they are never disturbed.
/// </summary>
public sealed class Deck
{
    private readonly OpcPackage _pkg;
    private readonly List<Slide> _slides = new();
    private Dialect _dialect = Dialect.Transitional;

    /// <summary>The family of names this deck was written in.</summary>
    public Dialect Dialect => _dialect;
    private Dialect D => _dialect;

    public OpcPackage Package => _pkg;
    public IReadOnlyList<Slide> Slides => _slides;

    public static bool Looks(OpcPackage pkg) => pkg.Has("ppt/presentation.xml");

    public Deck(OpcPackage pkg)
    {
        _pkg = pkg;
        if (!pkg.Has("ppt/presentation.xml")) throw new OpcPackage.PackageException("This is not a PowerPoint deck.");
        var pres = Xml.Parse(pkg.Read("ppt/presentation.xml"));
        _dialect = Dialect.Of(pres.Root!);
        var rels = new Rels(pkg, "ppt/presentation.xml");
        int number = 1;
        foreach (var id in pres.Root!.Element(D.Pres + "sldIdLst")?.Elements(D.Pres + "sldId") ?? Enumerable.Empty<XElement>())
        {
            string? rid = (string?)id.Attribute(D.Rel + "id");
            string? target = rid is null ? null : rels[rid];
            if (target is not null && pkg.Has(target)) _slides.Add(new Slide(this, number++, target));
        }
    }

    public void Flush() { foreach (var s in _slides) s.Flush(); }
    public void Save(string path) { Flush(); _pkg.Save(path); }
}

public sealed class Slide
{
    private readonly Deck _deck;
    private XDocument? _doc;
    private List<XElement>? _shapes;
    private bool _dirty;

    public int Number { get; }
    public string PartName { get; }

    private Dialect D => _deck.Dialect;

    internal Slide(Deck deck, int number, string partName) { _deck = deck; Number = number; PartName = partName; }

    private List<XElement> Shapes
    {
        get
        {
            if (_shapes is not null) return _shapes;
            _doc = Xml.Parse(_deck.Package.Read(PartName));
            _shapes = _doc.Root!.Descendants(D.Pres + "sp").Where(sp => Body(sp) is not null).ToList();
            return _shapes;
        }
    }

    /// <summary>A shape's text frame. Slides spell it p:txBody; drawing shapes elsewhere spell it a:txBody.</summary>
    private static XElement? Body(XElement shape) =>
        shape.Descendants().FirstOrDefault(e => e.Name.LocalName == "txBody");

    /// <summary>Every text frame on the slide, in the order the slide stores them.</summary>
    public IEnumerable<SlideText> Texts()
    {
        for (int i = 0; i < Shapes.Count; i++)
        {
            var sp = Shapes[i];
            // A shape with no placeholder element is plain text on the slide, not a body placeholder.
            var ph = sp.Descendants(D.Pres + "ph").FirstOrDefault();
            string placeholder = ph is null ? "" : (string?)ph.Attribute("type") ?? "body";
            var body = Body(sp)!;
            var shape = new TextShape(body, D.Draw);
            yield return new SlideText(i, placeholder, shape.Paragraphs.Select(shape.TextOf).ToList());
        }
    }

    /// <summary>The slide's title, which is what a thumbnail rail shows.</summary>
    public string Title() =>
        Texts().FirstOrDefault(t => t.Placeholder is "title" or "ctrTitle")?.Lines.FirstOrDefault(l => l.Length > 0)
        ?? Texts().SelectMany(t => t.Lines).FirstOrDefault(l => l.Length > 0)
        ?? "";

    /// <summary>Replace one line of one text frame. Returns false when mixed formatting in that line was flattened.</summary>
    public bool SetLine(int shapeIndex, int lineIndex, string value)
    {
        var body = Body(Shapes[shapeIndex])!;
        var shape = new TextShape(body, D.Draw);
        var paragraphs = shape.Paragraphs.ToList();
        if (lineIndex < 0 || lineIndex >= paragraphs.Count)
            throw new ArgumentOutOfRangeException(nameof(lineIndex), "That line is not on this slide.");
        bool lossless = shape.SetText(paragraphs[lineIndex], value);
        _dirty = true;
        return lossless;
    }

    internal void Flush()
    {
        if (!_dirty || _doc is null) return;
        _deck.Package.Write(PartName, Xml.ToBytes(_doc));
        _dirty = false;
    }
}
