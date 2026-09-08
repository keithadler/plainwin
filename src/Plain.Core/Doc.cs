using System.Xml.Linq;

namespace Plain.Core;

public enum BlockKind { Paragraph, Heading1, Heading2, Heading3, ListItem, TableCell }

/// <summary>
/// One block of a document as Plain shows it: the text, what kind of block it is, and where it lives. A block inside
/// a table also carries which table, row and column it sits in, so a table can be shown as a table rather than as a
/// column of loose cells.
/// </summary>
public sealed record Block(int Index, BlockKind Kind, string Text, bool Lossless,
                           int Table = -1, int Row = -1, int Column = -1)
{
    public bool IsHeading => Kind is BlockKind.Heading1 or BlockKind.Heading2 or BlockKind.Heading3;
    public bool InTable => Table >= 0;
}

/// <summary>
/// A Word document, opened through <see cref="OpcPackage"/>. Plain models the text of the body and nothing else, so
/// headers, footnotes, tracked changes, embedded objects and everything else stay in the file exactly as found.
/// </summary>
public sealed class Document
{
    public const string BodyPart = "word/document.xml";

    private readonly OpcPackage _pkg;
    private readonly XDocument _doc;
    private readonly TextShape _shape;
    private readonly List<XElement> _paragraphs;
    private bool _dirty;

    public OpcPackage Package => _pkg;
    public int BlockCount => _paragraphs.Count;

    public static bool Looks(OpcPackage pkg) => pkg.Has(BodyPart);

    public Document(OpcPackage pkg)
    {
        _pkg = pkg;
        if (!pkg.Has(BodyPart)) throw new OpcPackage.PackageException("This is not a Word document.");
        _doc = Xml.Parse(pkg.Read(BodyPart));
        _shape = new TextShape(_doc.Root!, Ns.Word);
        var body = _doc.Root!.Element(Ns.Word + "body") ?? _doc.Root!;
        _paragraphs = body.Descendants(Ns.Word + "p").ToList();

        // Work out each paragraph's place in a table once, rather than walking ancestors for every read.
        var tables = body.Descendants(Ns.Word + "tbl").ToList();
        var tableIndex = new Dictionary<XElement, int>();
        for (int i = 0; i < tables.Count; i++) tableIndex[tables[i]] = i;

        foreach (var p in _paragraphs)
        {
            var cell = p.Ancestors(Ns.Word + "tc").FirstOrDefault();
            var row = cell?.Parent;
            var table = row?.Parent;
            if (cell is null || row is null || table is null || !tableIndex.TryGetValue(table, out var t)) continue;
            int r = table.Elements(Ns.Word + "tr").ToList().IndexOf(row);
            int c = row.Elements(Ns.Word + "tc").ToList().IndexOf(cell);
            _place[p] = (t, r, c);
        }
    }

    private readonly Dictionary<XElement, (int Table, int Row, int Column)> _place = new();

    public IEnumerable<Block> Blocks()
    {
        for (int i = 0; i < _paragraphs.Count; i++) yield return Read(i);
    }

    public Block Read(int index)
    {
        var p = _paragraphs[index];
        string style = (string?)p.Element(Ns.Word + "pPr")?.Element(Ns.Word + "pStyle")?.Attribute(Ns.Word + "val") ?? "";
        bool inTable = p.Ancestors(Ns.Word + "tbl").Any();
        bool numbered = p.Element(Ns.Word + "pPr")?.Element(Ns.Word + "numPr") is not null;

        var kind = inTable ? BlockKind.TableCell
            : numbered ? BlockKind.ListItem
            : style.Contains("Heading1", StringComparison.OrdinalIgnoreCase) || style is "Title" ? BlockKind.Heading1
            : style.Contains("Heading2", StringComparison.OrdinalIgnoreCase) ? BlockKind.Heading2
            : style.Contains("Heading3", StringComparison.OrdinalIgnoreCase) ? BlockKind.Heading3
            : BlockKind.Paragraph;

        var place = _place.TryGetValue(p, out var found) ? found : (-1, -1, -1);
        return new Block(index, kind, _shape.TextOf(p), _shape.IsSingleRun(p) || _shape.TextOf(p).Length == 0,
                         place.Item1, place.Item2, place.Item3);
    }

    /// <summary>Replace a block's text. Returns false when mixed formatting inside that block was flattened.</summary>
    public bool SetText(int index, string value)
    {
        bool lossless = _shape.SetText(_paragraphs[index], value);
        _dirty = true;
        return lossless;
    }

    public string PlainText() => string.Join(Environment.NewLine, Blocks().Select(b => b.Text));

    public void Flush() { if (_dirty) { _pkg.Write(BodyPart, Xml.ToBytes(_doc)); _dirty = false; } }
    public void Save(string path) { Flush(); _pkg.Save(path); }
}
