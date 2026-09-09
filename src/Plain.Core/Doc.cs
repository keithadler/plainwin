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
    private List<XElement> _paragraphs = new();
    private readonly Dialect D;
    private bool _dirty;

    public OpcPackage Package => _pkg;
    public int BlockCount => _paragraphs.Count;

    public static bool Looks(OpcPackage pkg) => pkg.Has(BodyPart);

    public Document(OpcPackage pkg)
    {
        _pkg = pkg;
        if (!pkg.Has(BodyPart)) throw new OpcPackage.PackageException("This is not a Word document.");
        _doc = Xml.Parse(pkg.Read(BodyPart));
        D = Dialect.Of(_doc.Root!);
        _shape = new TextShape(_doc.Root!, D.Word);
        Index();
    }

    /// <summary>
    /// Work out where every paragraph is. Done when the file opens, and again whenever a row is added to or taken
    /// out of a table, because after that the old numbering points at the wrong paragraphs.
    /// </summary>
    private void Index()
    {
        var body = _doc.Root!.Element(D.Word + "body") ?? _doc.Root!;
        _paragraphs = body.Descendants(D.Word + "p").ToList();
        _place.Clear();

        // A document with no paragraphs at all means the body was not understood; say so rather than show a blank page.
        if (_paragraphs.Count == 0 && body.Elements().Any())
            throw new OpcPackage.PackageException(
                "Plain could not find the text in this document. It opens, but not in a way Plain reads.");

        // Work out each paragraph's place in a table once, rather than walking ancestors for every read.
        var tables = body.Descendants(D.Word + "tbl").ToList();
        var tableIndex = new Dictionary<XElement, int>();
        for (int i = 0; i < tables.Count; i++) tableIndex[tables[i]] = i;

        foreach (var p in _paragraphs)
        {
            var cell = p.Ancestors(D.Word + "tc").FirstOrDefault();
            var row = cell?.Parent;
            var table = row?.Parent;
            if (cell is null || row is null || table is null || !tableIndex.TryGetValue(table, out var t)) continue;
            int r = table.Elements(D.Word + "tr").ToList().IndexOf(row);
            int c = row.Elements(D.Word + "tc").ToList().IndexOf(cell);
            _place[p] = (t, r, c);
        }
    }

    private readonly Dictionary<XElement, (int Table, int Row, int Column)> _place = new();

    // ---------- the shape of a table ----------

    /// <summary>How many rows each table in this document has, in the order they appear.</summary>
    public IReadOnlyList<int> TableShape()
    {
        var body = _doc.Root!.Element(D.Word + "body") ?? _doc.Root!;
        return body.Descendants(D.Word + "tbl").Select(t => t.Elements(D.Word + "tr").Count()).ToList();
    }

    private List<XElement>? RowsOf(int table)
    {
        var body = _doc.Root!.Element(D.Word + "body") ?? _doc.Root!;
        var tables = body.Descendants(D.Word + "tbl").ToList();
        if (table < 0 || table >= tables.Count) return null;
        return tables[table].Elements(D.Word + "tr").ToList();
    }

    /// <summary>
    /// Put a new empty row into a table, after the row given, or at the top when that is nought. The new row is a
    /// copy of a neighbour with its words taken out, so it keeps the borders, shading and widths the table uses.
    /// Building one from scratch would produce a row that looked nothing like the table it joined.
    /// </summary>
    /// <summary>
    /// Put a new empty paragraph in after the one given, taking its shape from the paragraph it follows.
    ///
    /// A new document has exactly one paragraph in it, and without this it has exactly one for ever: you could
    /// type into it and nowhere else. The new one copies its neighbour's properties and none of its words, so a
    /// paragraph added after a heading is a heading and one added after a bullet is a bullet, which is what
    /// pressing Return in a word processor does.
    ///
    /// It refuses inside a table, where a paragraph is a cell and adding one changes the shape of the table
    /// rather than the text.
    /// </summary>
    public TableRows.Result InsertParagraph(int after)
    {
        if (_paragraphs.Count == 0) return new TableRows.Refused("This document has no paragraphs to add to.");

        int at = Math.Clamp(after, -1, _paragraphs.Count - 1);
        var neighbour = _paragraphs[Math.Max(0, at)];
        if (neighbour.Ancestors(D.Word + "tbl").Any())
            return new TableRows.Refused("That paragraph is in a table. Add a row to the table instead.");

        var fresh = new XElement(neighbour);
        // The shape without the words: properties stay, everything that carries text goes.
        foreach (var name in new[] { "r", "ins", "del", "hyperlink", "bookmarkStart", "bookmarkEnd", "commentRangeStart", "commentRangeEnd" })
            fresh.Elements(D.Word + name).Remove();

        if (after < 0) neighbour.AddBeforeSelf(fresh); else neighbour.AddAfterSelf(fresh);

        _dirty = true;
        Index();
        return new TableRows.Done($"A paragraph is in, making {_paragraphs.Count}.");
    }

    /// <summary>
    /// Take a paragraph out. The last one cannot go: a document with no paragraphs at all is one Word will not
    /// open, and leaving somebody with a file they cannot reopen is worse than refusing.
    /// </summary>
    public TableRows.Result DeleteParagraph(int index)
    {
        if (index < 0 || index >= _paragraphs.Count) return new TableRows.Refused($"There is no paragraph {index + 1}.");
        if (_paragraphs.Count <= 1) return new TableRows.Refused("A document has to have a paragraph in it.");

        var paragraph = _paragraphs[index];
        if (paragraph.Ancestors(D.Word + "tbl").Any())
            return new TableRows.Refused("That paragraph is in a table. Take out a row of the table instead.");

        paragraph.Remove();
        _dirty = true;
        Index();
        return new TableRows.Done($"That paragraph is out, leaving {_paragraphs.Count}.");
    }

    public TableRows.Result InsertRow(int table, int after)
    {
        var rows = RowsOf(table);
        if (rows is null) return new TableRows.Refused($"There is no table {table + 1} in this document.");
        if (rows.Count == 0) return new TableRows.Refused("That table has no rows to copy the shape of.");

        int copyFrom = Math.Clamp(after - 1, 0, rows.Count - 1);
        var fresh = new XElement(rows[copyFrom]);
        foreach (var paragraph in fresh.Descendants(D.Word + "p").ToList())
        {
            foreach (var run in paragraph.Elements(D.Word + "r").ToList()) run.Remove();
            foreach (var inserted in paragraph.Elements(D.Word + "ins").ToList()) inserted.Remove();
            foreach (var deleted in paragraph.Elements(D.Word + "del").ToList()) deleted.Remove();
        }

        if (after <= 0) rows[0].AddBeforeSelf(fresh);
        else if (after >= rows.Count) rows[^1].AddAfterSelf(fresh);
        else rows[after - 1].AddAfterSelf(fresh);

        _dirty = true;
        Index();
        return new TableRows.Done($"A row is in, making {rows.Count + 1}.");
    }

    /// <summary>Take a row out of a table. The last row of a table cannot go, because Word calls that damage.</summary>
    public TableRows.Result DeleteRow(int table, int row)
    {
        var rows = RowsOf(table);
        if (rows is null) return new TableRows.Refused($"There is no table {table + 1} in this document.");
        if (row < 1 || row > rows.Count) return new TableRows.Refused($"That table has no row {row}.");
        if (rows.Count <= 1) return new TableRows.Refused("A table has to have a row in it, so this one cannot go.");

        rows[row - 1].Remove();
        _dirty = true;
        Index();
        return new TableRows.Done($"That row is out, leaving {rows.Count - 1}.");
    }

    public IEnumerable<Block> Blocks()
    {
        for (int i = 0; i < _paragraphs.Count; i++) yield return Read(i);
    }

    public Block Read(int index)
    {
        var p = _paragraphs[index];
        string style = (string?)p.Element(D.Word + "pPr")?.Element(D.Word + "pStyle")?.Attribute(D.Word + "val") ?? "";
        bool inTable = p.Ancestors(D.Word + "tbl").Any();
        var properties = p.Element(D.Word + "pPr");

        // A paragraph can carry numbering that says "none". numId 0 means exactly that, and treating a
        // paragraph that has one as a list item turns every heading LibreOffice writes into a bullet.
        var numbering = properties?.Element(D.Word + "numPr");
        bool numbered = numbering is not null
                     && Xml.Int(numbering.Element(D.Word + "numId")?.Attribute(D.Word + "val"), 1) != 0;

        // Word marks a heading by its style, but a heading can also say its level outright, which is what
        // LibreOffice writes when it exports one. Reading only the style name misses those entirely.
        int outline = Xml.Int(properties?.Element(D.Word + "outlineLvl")?.Attribute(D.Word + "val"), -1);

        var kind = inTable ? BlockKind.TableCell
            : style.Contains("Heading1", StringComparison.OrdinalIgnoreCase) || style is "Title" ? BlockKind.Heading1
            : style.Contains("Heading2", StringComparison.OrdinalIgnoreCase) ? BlockKind.Heading2
            : style.Contains("Heading3", StringComparison.OrdinalIgnoreCase) ? BlockKind.Heading3
            : outline == 0 ? BlockKind.Heading1
            : outline == 1 ? BlockKind.Heading2
            : outline >= 2 ? BlockKind.Heading3
            : numbered ? BlockKind.ListItem
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

    /// <summary>Is every run in this block already bold, or italic? Used to make the button show what is true.</summary>
    public bool IsAll(int index, string mark)
    {
        var runs = _paragraphs[index].Elements(D.Word + "r").ToList();
        if (runs.Count == 0) return false;
        return runs.All(r => r.Element(D.Word + "rPr")?.Element(D.Word + mark) is not null);
    }

    /// <summary>
    /// Turn bold or italic on or off for a whole block. Plain edits a block at a time, so it formats one at a time
    /// too: enough for a heading that looks like a heading and a notice that stands out, which is what people asked
    /// for, without pretending to be a word processor.
    /// </summary>
    public void SetMark(int index, string mark, bool on)
    {
        foreach (var run in _paragraphs[index].Elements(D.Word + "r"))
        {
            var properties = run.Element(D.Word + "rPr");
            if (on)
            {
                if (properties is null) { properties = new XElement(D.Word + "rPr"); run.AddFirst(properties); }
                if (properties.Element(D.Word + mark) is null) properties.AddFirst(new XElement(D.Word + mark));
            }
            else properties?.Elements(D.Word + mark).Remove();
        }
        _dirty = true;
    }

    /// <summary>
    /// Make a block a heading, or ordinary text again. The style has to exist in the document already; Plain does not
    /// invent one, because a heading style it made up would not match the rest of somebody's document.
    /// </summary>
    public bool SetKind(int index, BlockKind kind)
    {
        string? style = kind switch
        {
            BlockKind.Heading1 => "Heading1",
            BlockKind.Heading2 => "Heading2",
            BlockKind.Heading3 => "Heading3",
            BlockKind.Paragraph => null,
            _ => throw new ArgumentOutOfRangeException(nameof(kind), "Plain sets headings and ordinary text."),
        };

        if (style is not null && !HasStyle(style)) return false;

        var paragraph = _paragraphs[index];
        var properties = paragraph.Element(D.Word + "pPr");
        if (style is null)
        {
            properties?.Elements(D.Word + "pStyle").Remove();
        }
        else
        {
            if (properties is null) { properties = new XElement(D.Word + "pPr"); paragraph.AddFirst(properties); }
            properties.Elements(D.Word + "pStyle").Remove();
            properties.AddFirst(new XElement(D.Word + "pStyle", new XAttribute(D.Word + "val", style)));
        }
        _dirty = true;
        return true;
    }

    /// <summary>
    /// Where a paragraph sits across the page: left, centred, right or justified. "general" takes it back to
    /// whatever its style says, which is not the same as setting it to left, and is what someone means when they
    /// want to undo having centred something.
    /// </summary>
    public bool SetParagraphAlignment(int index, string where)
    {
        if (index < 0 || index >= _paragraphs.Count) return false;

        string? value = where.ToLowerInvariant() switch
        {
            "left" => "left",
            "centre" or "center" => "center",
            "right" => "right",
            "justify" or "justified" => "both",
            "general" or "none" => null,
            _ => "?",
        };
        if (value == "?") return false;

        var paragraph = _paragraphs[index];
        var properties = paragraph.Element(D.Word + "pPr");

        if (value is null)
        {
            properties?.Elements(D.Word + "jc").Remove();
            _dirty = true;
            return true;
        }

        if (properties is null) { properties = new XElement(D.Word + "pPr"); paragraph.AddFirst(properties); }
        properties.Elements(D.Word + "jc").Remove();
        // jc comes after pStyle and the numbering, which is the order Word writes and expects.
        var after = properties.Element(D.Word + "numPr") ?? properties.Element(D.Word + "pStyle");
        var element = new XElement(D.Word + "jc", new XAttribute(D.Word + "val", value));
        if (after is not null) after.AddAfterSelf(element); else properties.AddFirst(element);

        _dirty = true;
        return true;
    }

    /// <summary>Where this paragraph sits, as the word someone would use for it.</summary>
    public string ParagraphAlignment(int index)
    {
        if (index < 0 || index >= _paragraphs.Count) return "general";
        var value = (string?)_paragraphs[index].Element(D.Word + "pPr")?.Element(D.Word + "jc")?.Attribute(D.Word + "val");
        return value switch
        {
            "center" => "centre",
            "right" => "right",
            "both" or "distribute" => "justify",
            "left" or "start" => "left",
            _ => "general",
        };
    }

    private bool HasStyle(string id)
    {
        if (!_pkg.Has("word/styles.xml")) return false;
        try
        {
            var styles = Xml.Parse(_pkg.Read("word/styles.xml"));
            var d = Dialect.Of(styles.Root!);
            return styles.Root!.Elements(d.Word + "style")
                .Any(x => string.Equals((string?)x.Attribute(d.Word + "styleId"), id, StringComparison.OrdinalIgnoreCase));
        }
        catch { return false; }
    }

    /// <summary>
    /// Make a block a bullet or a numbered item, or take it out of the list. Word keeps list numbering in its own
    /// part, so Plain uses a definition the document already has rather than inventing one; a document with no list
    /// definitions is told so instead of being given a broken list.
    /// </summary>
    public bool SetList(int index, bool bulleted, bool on)
    {
        var paragraph = _paragraphs[index];
        var properties = paragraph.Element(D.Word + "pPr");

        if (!on)
        {
            properties?.Elements(D.Word + "numPr").Remove();
            _dirty = true;
            return true;
        }

        int? numbering = ListDefinition(bulleted);
        if (numbering is not { } id) return false;

        if (properties is null) { properties = new XElement(D.Word + "pPr"); paragraph.AddFirst(properties); }
        properties.Elements(D.Word + "numPr").Remove();
        var marker = new XElement(D.Word + "numPr",
            new XElement(D.Word + "ilvl", new XAttribute(D.Word + "val", 0)),
            new XElement(D.Word + "numId", new XAttribute(D.Word + "val", id)));
        // The properties element has an order; numPr goes after the style if there is one.
        var style = properties.Element(D.Word + "pStyle");
        if (style is not null) style.AddAfterSelf(marker); else properties.AddFirst(marker);
        _dirty = true;
        return true;
    }

    public bool IsList(int index) =>
        _paragraphs[index].Element(D.Word + "pPr")?.Element(D.Word + "numPr") is not null;

    /// <summary>A list definition of the wanted shape that this document already carries, if it has one.</summary>
    private int? ListDefinition(bool bulleted)
    {
        if (!_pkg.Has("word/numbering.xml")) return null;
        try
        {
            var numbering = Xml.Parse(_pkg.Read("word/numbering.xml"));
            var d = Dialect.Of(numbering.Root!);

            var abstracts = new Dictionary<string, bool>();   // abstract id -> is it a bullet list
            foreach (var definition in numbering.Root!.Elements(d.Word + "abstractNum"))
            {
                var id = (string?)definition.Attribute(d.Word + "abstractNumId");
                if (id is null) continue;
                var level = definition.Elements(d.Word + "lvl")
                    .FirstOrDefault(l => (string?)l.Attribute(d.Word + "ilvl") == "0");
                var format = (string?)level?.Element(d.Word + "numFmt")?.Attribute(d.Word + "val") ?? "";
                abstracts[id] = format.Equals("bullet", StringComparison.OrdinalIgnoreCase);
            }

            foreach (var use in numbering.Root.Elements(d.Word + "num"))
            {
                var id = (string?)use.Attribute(d.Word + "numId");
                var pointsAt = (string?)use.Element(d.Word + "abstractNumId")?.Attribute(d.Word + "val");
                if (id is null || pointsAt is null) continue;
                if (!abstracts.TryGetValue(pointsAt, out var isBullet) || isBullet != bulleted) continue;
                if (int.TryParse(id, out var number)) return number;
            }
        }
        catch { }
        return null;
    }

    /// <summary>Whether this document can do bullets, numbers, both or neither.</summary>
    public (bool Bullets, bool Numbers) ListsAvailable() => (ListDefinition(true) is not null, ListDefinition(false) is not null);

    /// <summary>Which heading styles this document actually has, so only those are offered.</summary>
    public IReadOnlyList<BlockKind> AvailableKinds()
    {
        var kinds = new List<BlockKind> { BlockKind.Paragraph };
        if (HasStyle("Heading1")) kinds.Add(BlockKind.Heading1);
        if (HasStyle("Heading2")) kinds.Add(BlockKind.Heading2);
        if (HasStyle("Heading3")) kinds.Add(BlockKind.Heading3);
        return kinds;
    }

    public string PlainText() => string.Join(Environment.NewLine, Blocks().Select(b => b.Text));

    public void Flush() { if (_dirty) { _pkg.Write(BodyPart, Xml.ToBytes(_doc)); _dirty = false; } }
    public void Save(string path) { Flush(); _pkg.Save(path); }
}
