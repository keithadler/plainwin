using System.Globalization;
using System.Xml.Linq;

namespace Plain.Core;

public enum CellKind { Empty, Number, Text, Formula, Boolean, Error }

/// <summary>One cell as Plain shows it: what it holds, what Excel last computed, and how it reads on screen.</summary>
public sealed record Cell(CellRef Ref, CellKind Kind, string Raw, string Display, string? Formula)
{
    public bool IsEmpty => Kind == CellKind.Empty;
}

/// <summary>
/// An Excel workbook, opened through <see cref="OpcPackage"/> so everything Plain does not model - charts, pivot
/// tables, conditional formatting, macros - stays in the file untouched.
/// </summary>
public sealed class Workbook
{
    private readonly OpcPackage _pkg;
    private readonly List<Sheet> _sheets = new();
    private XDocument? _sst;             // xl/sharedStrings.xml, loaded on demand
    private List<string>? _strings;
    private bool _sstDirty;
    private XDocument? _workbookDoc;
    private bool _workbookDirty;
    private Styles? _styles;

    public OpcPackage Package => _pkg;
    public IReadOnlyList<Sheet> Sheets => _sheets;
    public string WorkbookPart { get; }

    public static bool Looks(OpcPackage pkg) => pkg.Has("xl/workbook.xml");

    public Workbook(OpcPackage pkg)
    {
        _pkg = pkg;
        WorkbookPart = "xl/workbook.xml";
        if (!pkg.Has(WorkbookPart)) throw new OpcPackage.PackageException("This is not an Excel workbook.");

        _workbookDoc = Xml.Parse(pkg.Read(WorkbookPart));
        var rels = new Rels(pkg, WorkbookPart);
        foreach (var s in _workbookDoc.Root!.Element(Ns.Sheet + "sheets")?.Elements(Ns.Sheet + "sheet") ?? Enumerable.Empty<XElement>())
        {
            string name = (string?)s.Attribute("name") ?? "Sheet";
            string? rid = (string?)s.Attribute(Ns.Rel + "id");
            string? target = rid is null ? null : rels[rid];
            bool hidden = ((string?)s.Attribute("state") ?? "visible") != "visible";
            if (target is not null && pkg.Has(target)) _sheets.Add(new Sheet(this, name, target, hidden));
        }
    }

    internal Styles Styles => _styles ??= new Styles(_pkg);

    // ---------- shared strings ----------

    internal string SharedString(int index)
    {
        LoadStrings();
        return index >= 0 && index < _strings!.Count ? _strings[index] : "";
    }

    private void LoadStrings()
    {
        if (_strings is not null) return;
        _strings = new List<string>();
        if (!_pkg.Has("xl/sharedStrings.xml")) return;
        _sst = Xml.Parse(_pkg.Read("xl/sharedStrings.xml"));
        foreach (var si in _sst.Root!.Elements(Ns.Sheet + "si")) _strings.Add(SiText(si));
    }

    private static string SiText(XElement si)
    {
        // A shared string is either one <t> or a run of <r><t> pieces with their own formatting.
        var t = si.Element(Ns.Sheet + "t");
        if (t is not null) return t.Value;
        return string.Concat(si.Elements(Ns.Sheet + "r").Select(r => r.Element(Ns.Sheet + "t")?.Value ?? ""));
    }

    /// <summary>The index for a piece of text, appending it to the shared table when it is new.</summary>
    internal int InternString(string text)
    {
        LoadStrings();
        int existing = _strings!.IndexOf(text);
        if (existing >= 0) return existing;
        if (_sst is null) throw new OpcPackage.PackageException("This workbook has no shared string table, so Plain cannot add text to it yet.");

        var si = new XElement(Ns.Sheet + "si", new XElement(Ns.Sheet + "t", text));
        // Text with leading or trailing spaces needs the xml:space hint or Excel trims it.
        if (text.Length > 0 && (char.IsWhiteSpace(text[0]) || char.IsWhiteSpace(text[^1])))
            si.Element(Ns.Sheet + "t")!.SetAttributeValue(XNamespace.Xml + "space", "preserve");
        _sst.Root!.Add(si);
        _strings.Add(text);
        _sst.Root.SetAttributeValue("uniqueCount", _strings.Count.ToString(CultureInfo.InvariantCulture));
        int count = int.TryParse((string?)_sst.Root.Attribute("count"), NumberStyles.Integer, CultureInfo.InvariantCulture, out var c) ? c : _strings.Count;
        _sst.Root.SetAttributeValue("count", (count + 1).ToString(CultureInfo.InvariantCulture));
        _sstDirty = true;
        return _strings.Count - 1;
    }

    /// <summary>
    /// Tell Excel to recalculate everything when it opens the file. Plain does not evaluate formulas, so a cached
    /// value next to an edited cell would otherwise be quietly stale - the one thing a spreadsheet must never be.
    /// </summary>
    internal void RequestFullRecalculation()
    {
        var root = _workbookDoc!.Root!;
        var calcPr = root.Element(Ns.Sheet + "calcPr");
        if (calcPr is null)
        {
            calcPr = new XElement(Ns.Sheet + "calcPr", new XAttribute("calcId", "0"));
            var sheets = root.Element(Ns.Sheet + "sheets");
            if (sheets is not null) sheets.AddAfterSelf(calcPr); else root.Add(calcPr);
        }
        calcPr.SetAttributeValue("fullCalcOnLoad", "1");
        _workbookDirty = true;
    }

    /// <summary>Write every part Plain changed back into the package. Nothing else is touched.</summary>
    public void Flush()
    {
        foreach (var sheet in _sheets) sheet.Flush();
        if (_sstDirty && _sst is not null) { _pkg.Write("xl/sharedStrings.xml", Xml.ToBytes(_sst)); _sstDirty = false; }
        if (_workbookDirty && _workbookDoc is not null) { _pkg.Write(WorkbookPart, Xml.ToBytes(_workbookDoc)); _workbookDirty = false; }
    }

    public void Save(string path) { Flush(); _pkg.Save(path); }
}

/// <summary>One worksheet. Cells are read straight from the sheet's XML and edited in place.</summary>
public sealed class Sheet
{
    private readonly Workbook _book;
    private XDocument? _doc;
    private XElement? _data;
    private bool _dirty;

    public string Name { get; }
    public string PartName { get; }
    public bool Hidden { get; }

    internal Sheet(Workbook book, string name, string partName, bool hidden)
    {
        _book = book; Name = name; PartName = partName; Hidden = hidden;
    }

    private XElement Data
    {
        get
        {
            if (_data is not null) return _data;
            _widths = null;
            _doc = Xml.Parse(_book.Package.Read(PartName));
            _data = _doc.Root!.Element(Ns.Sheet + "sheetData")
                    ?? throw new OpcPackage.PackageException($"The sheet \"{Name}\" has no cell data.");
            return _data;
        }
    }

    /// <summary>The furthest cell that holds anything, which is how wide and tall Plain draws the grid.</summary>
    public CellRef Extent
    {
        get
        {
            int maxCol = 1, maxRow = 1;
            foreach (var row in Data.Elements(Ns.Sheet + "row"))
                foreach (var c in row.Elements(Ns.Sheet + "c"))
                    if (CellRef.TryParse((string?)c.Attribute("r") ?? "", out var r) && !string.IsNullOrEmpty(c.Value))
                    {
                        if (r.Column > maxCol) maxCol = r.Column;
                        if (r.Row > maxRow) maxRow = r.Row;
                    }
            return new CellRef(maxCol, maxRow);
        }
    }

    /// <summary>
    /// The column widths the file stores, in Excel's character units. Drawing a grid at the widths someone chose is
    /// most of what makes a spreadsheet look like theirs rather than like a generic table.
    /// </summary>
    public double WidthChars(int column)
    {
        _widths ??= ReadWidths();
        foreach (var (min, max, width) in _widths) if (column >= min && column <= max) return width;
        return _defaultWidth;
    }

    private List<(int Min, int Max, double Width)>? _widths;
    private double _defaultWidth = 8.43;

    private List<(int, int, double)> ReadWidths()
    {
        var list = new List<(int, int, double)>();
        var sheetElement = _doc?.Root ?? Xml.Parse(_book.Package.Read(PartName)).Root;
        var format = sheetElement?.Element(Ns.Sheet + "sheetFormatPr");
        if (format is not null && double.TryParse((string?)format.Attribute("defaultColWidth"),
                System.Globalization.NumberStyles.Float, CultureInfo.InvariantCulture, out var d) && d > 0)
            _defaultWidth = d;
        foreach (var c in sheetElement?.Element(Ns.Sheet + "cols")?.Elements(Ns.Sheet + "col") ?? Enumerable.Empty<XElement>())
        {
            if (!int.TryParse((string?)c.Attribute("min"), out var min)) continue;
            if (!int.TryParse((string?)c.Attribute("max"), out var max)) continue;
            if (!double.TryParse((string?)c.Attribute("width"), System.Globalization.NumberStyles.Float, CultureInfo.InvariantCulture, out var w)) continue;
            list.Add((min, Math.Min(max, 16384), w));
        }
        return list;
    }

    private XElement? FindRow(int row) =>
        Data.Elements(Ns.Sheet + "row").FirstOrDefault(r => (int?)r.Attribute("r") == row);

    private XElement? FindCell(CellRef cell)
    {
        string want = cell.ToString();
        return FindRow(cell.Row)?.Elements(Ns.Sheet + "c").FirstOrDefault(c => (string?)c.Attribute("r") == want);
    }

    public Cell Read(string reference) => Read(CellRef.Parse(reference));

    public Cell Read(CellRef reference)
    {
        var c = FindCell(reference);
        if (c is null) return new Cell(reference, CellKind.Empty, "", "", null);

        string? type = (string?)c.Attribute("t");
        string? formula = c.Element(Ns.Sheet + "f")?.Value;
        var v = c.Element(Ns.Sheet + "v");

        string raw, display;
        CellKind kind;
        switch (type)
        {
            case "s":
                raw = display = int.TryParse(v?.Value, NumberStyles.Integer, CultureInfo.InvariantCulture, out var idx) ? _book.SharedString(idx) : "";
                kind = CellKind.Text;
                break;
            case "inlineStr":
                raw = display = c.Element(Ns.Sheet + "is")?.Value ?? "";
                kind = CellKind.Text;
                break;
            case "str":
                raw = display = v?.Value ?? "";
                kind = CellKind.Text;
                break;
            case "b":
                raw = v?.Value ?? "0";
                display = raw == "1" ? "TRUE" : "FALSE";
                kind = CellKind.Boolean;
                break;
            case "e":
                raw = display = v?.Value ?? "#N/A";
                kind = CellKind.Error;
                break;
            default:
                raw = v?.Value ?? "";
                kind = raw.Length == 0 ? CellKind.Empty : CellKind.Number;
                display = _book.Styles.Format(raw, (int?)c.Attribute("s") ?? 0);
                break;
        }

        if (formula is not null)
        {
            kind = CellKind.Formula;
            // Plain clears the cached value beside a formula when the sheet changes, so there is often nothing to
            // show. Showing the formula itself beats showing an empty cell where a total used to be.
            if (display.Length == 0) display = "=" + formula;
        }
        return new Cell(reference, kind, raw, display, formula is null ? null : "=" + formula);
    }

    /// <summary>Every non-empty cell, in reading order.</summary>
    public IEnumerable<Cell> Cells()
    {
        foreach (var row in Data.Elements(Ns.Sheet + "row"))
            foreach (var c in row.Elements(Ns.Sheet + "c"))
                if (CellRef.TryParse((string?)c.Attribute("r") ?? "", out var r))
                {
                    var cell = Read(r);
                    if (!cell.IsEmpty || cell.Formula is not null) yield return cell;
                }
    }

    // ---------- editing ----------

    /// <summary>
    /// Set a cell from what someone typed: a leading "=" makes a formula, a plain number makes a number, anything
    /// else is text. This is the only editing Plain offers on a sheet, and it is the one people use.
    /// </summary>
    public void Set(CellRef reference, string typed)
    {
        var c = EnsureCell(reference);
        c.Elements(Ns.Sheet + "f").Remove();
        c.Elements(Ns.Sheet + "v").Remove();
        c.Elements(Ns.Sheet + "is").Remove();
        c.Attribute("t")?.Remove();

        if (typed.StartsWith('=') && typed.Length > 1)
        {
            c.Add(new XElement(Ns.Sheet + "f", typed[1..]));
            _book.RequestFullRecalculation();
        }
        else if (typed.Length == 0)
        {
            // An emptied cell keeps its formatting and loses its content, the way Delete behaves in Excel.
        }
        else if (double.TryParse(typed, NumberStyles.Float, CultureInfo.InvariantCulture, out var number)
                 && !typed.Contains(' ') && double.IsFinite(number))
        {
            c.Add(new XElement(Ns.Sheet + "v", number.ToString("R", CultureInfo.InvariantCulture)));
            _book.RequestFullRecalculation();
        }
        else
        {
            c.SetAttributeValue("t", "s");
            c.Add(new XElement(Ns.Sheet + "v", _book.InternString(typed).ToString(CultureInfo.InvariantCulture)));
            _book.RequestFullRecalculation();
        }
        _dirty = true;
    }

    public void Set(string reference, string typed) => Set(CellRef.Parse(reference), typed);

    private XElement EnsureCell(CellRef reference)
    {
        var row = FindRow(reference.Row);
        if (row is null)
        {
            row = new XElement(Ns.Sheet + "row", new XAttribute("r", reference.Row));
            var after = Data.Elements(Ns.Sheet + "row").LastOrDefault(r => ((int?)r.Attribute("r") ?? 0) < reference.Row);
            if (after is not null) after.AddAfterSelf(row); else Data.AddFirst(row);
        }

        string want = reference.ToString();
        var cell = row.Elements(Ns.Sheet + "c").FirstOrDefault(c => (string?)c.Attribute("r") == want);
        if (cell is not null) return cell;

        cell = new XElement(Ns.Sheet + "c", new XAttribute("r", want));
        // Cells must stay in ascending column order or Excel reports the file as damaged.
        var prev = row.Elements(Ns.Sheet + "c")
            .LastOrDefault(c => CellRef.TryParse((string?)c.Attribute("r") ?? "", out var r) && r.Column < reference.Column);
        if (prev is not null) prev.AddAfterSelf(cell); else row.AddFirst(cell);
        return cell;
    }

    /// <summary>
    /// Drop the value Excel last cached beside each formula on this sheet. Plain does not evaluate formulas, so once
    /// a cell on the sheet changes, every cached total near it is a number that may no longer be true. A formula with
    /// no cached value forces whatever opens the file to work it out, which is the only way to be sure the number on
    /// screen is the right one. Cells without a formula keep their values.
    /// </summary>
    private void DropStaleCachedValues()
    {
        foreach (var row in Data.Elements(Ns.Sheet + "row"))
            foreach (var c in row.Elements(Ns.Sheet + "c"))
                if (c.Element(Ns.Sheet + "f") is not null)
                {
                    c.Elements(Ns.Sheet + "v").Remove();
                    // A cached string result is typed on the cell; without the value the type is meaningless.
                    if ((string?)c.Attribute("t") == "str") c.Attribute("t")!.Remove();
                }
    }

    internal void Flush()
    {
        if (!_dirty || _doc is null) return;
        DropStaleCachedValues();
        _book.Package.Write(PartName, Xml.ToBytes(_doc));
        _dirty = false;
    }
}
