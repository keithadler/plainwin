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
    private Dialect _dialect = Dialect.Transitional;

    /// <summary>The family of names this workbook was written in, so every part is read the way it was written.</summary>
    public Dialect Dialect => _dialect;
    private Dialect D => _dialect;

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
        _dialect = Dialect.Of(_workbookDoc.Root!);
        var rels = new Rels(pkg, WorkbookPart);
        foreach (var s in _workbookDoc.Root!.Element(D.Sheet + "sheets")?.Elements(D.Sheet + "sheet") ?? Enumerable.Empty<XElement>())
        {
            string name = (string?)s.Attribute("name") ?? "Sheet";
            string? rid = (string?)s.Attribute(D.Rel + "id");
            string? target = rid is null ? null : rels[rid];
            bool hidden = ((string?)s.Attribute("state") ?? "visible") != "visible";
            if (target is not null && pkg.Has(target)) _sheets.Add(new Sheet(this, name, target, hidden));
        }

        // A workbook always has at least one sheet. Finding none means Plain did not understand the file, and
        // showing an empty grid would be a lie about what is in it.
        if (_sheets.Count == 0)
            throw new OpcPackage.PackageException(
                "Plain could not find the sheets in this workbook. It opens, but not in a way Plain reads, so nothing is shown rather than showing you an empty grid.");
    }

    internal Styles Styles => _styles ??= new Styles(_pkg, _dialect);

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
        foreach (var si in _sst.Root!.Elements(D.Sheet + "si")) _strings.Add(SiText(si));
    }

    private string SiText(XElement si)
    {
        // A shared string is either one <t> or a run of <r><t> pieces with their own formatting.
        var t = si.Element(D.Sheet + "t");
        if (t is not null) return t.Value;
        return string.Concat(si.Elements(D.Sheet + "r").Select(r => r.Element(D.Sheet + "t")?.Value ?? ""));
    }

    /// <summary>The index for a piece of text, appending it to the shared table when it is new.</summary>
    internal int InternString(string text)
    {
        LoadStrings();
        int existing = _strings!.IndexOf(text);
        if (existing >= 0) return existing;
        if (_sst is null) throw new OpcPackage.PackageException("This workbook has no shared string table, so Plain cannot add text to it yet.");

        var si = new XElement(D.Sheet + "si", new XElement(D.Sheet + "t", text));
        // Text with leading or trailing spaces needs the xml:space hint or Excel trims it.
        if (text.Length > 0 && (char.IsWhiteSpace(text[0]) || char.IsWhiteSpace(text[^1])))
            si.Element(D.Sheet + "t")!.SetAttributeValue(XNamespace.Xml + "space", "preserve");
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
        var calcPr = root.Element(D.Sheet + "calcPr");
        if (calcPr is null)
        {
            calcPr = new XElement(D.Sheet + "calcPr", new XAttribute("calcId", "0"));
            var sheets = root.Element(D.Sheet + "sheets");
            if (sheets is not null) sheets.AddAfterSelf(calcPr); else root.Add(calcPr);
        }
        calcPr.SetAttributeValue("fullCalcOnLoad", "1");
        _workbookDirty = true;
    }

    private readonly Dictionary<Sheet, HashSet<CellRef>> _changed = new();

    internal void NoteChanged(Sheet sheet, CellRef cell)
    {
        if (!_changed.TryGetValue(sheet, out var set)) _changed[sheet] = set = new HashSet<CellRef>();
        set.Add(cell);
    }

    /// <summary>
    /// Work out which cached totals an edit made untrue, and clear only those. A formula that reads an edited cell is
    /// stale, and so is a formula that reads that formula, all the way up the chain. Everything else keeps the value
    /// Excel worked out, so changing one label in a budget does not empty the budget of its numbers.
    /// </summary>
    private void ClearStaleCaches()
    {
        if (_changed.Count == 0) return;

        var byName = new Dictionary<string, int>(StringComparer.CurrentCultureIgnoreCase);
        for (int i = 0; i < _sheets.Count; i++) byName[_sheets[i].Name] = i;

        var dirty = new Dictionary<int, HashSet<CellRef>>();
        void MarkDirty(int sheet, CellRef cell)
        {
            if (!dirty.TryGetValue(sheet, out var set)) dirty[sheet] = set = new HashSet<CellRef>();
            set.Add(cell);
        }
        foreach (var (sheet, cells) in _changed)
        {
            int index = _sheets.IndexOf(sheet);
            foreach (var cell in cells) MarkDirty(index, cell);
        }

        // Read every formula once, with the cells it depends on already worked out.
        var formulas = new List<(int Sheet, CellRef Ref, RefRange[] Reads)>();
        for (int i = 0; i < _sheets.Count; i++)
            foreach (var (reference, formula) in _sheets[i].Formulas())
                formulas.Add((i, reference, Refs.Parse(formula).ToArray()));

        var stale = new HashSet<(int, CellRef)>();

        // A formula somebody has just typed needs working out even when it reads nothing that changed - typing
        // "=1+1" into a cell must show 2, not stay empty waiting for something else to move.
        foreach (var formula in formulas)
            if (dirty.TryGetValue(formula.Sheet, out var touched) && touched.Contains(formula.Ref))
                stale.Add((formula.Sheet, formula.Ref));
        // Each pass can only make more cells stale, so this settles in at most one pass per formula.
        for (int pass = 0; pass <= formulas.Count; pass++)
        {
            bool grew = false;
            foreach (var (sheet, reference, reads) in formulas)
            {
                if (stale.Contains((sheet, reference))) continue;
                if (!ReadsAnything(reads, sheet, byName, dirty)) continue;
                stale.Add((sheet, reference));
                MarkDirty(sheet, reference);
                grew = true;
            }
            if (!grew) break;
        }

        for (int i = 0; i < _sheets.Count; i++)
        {
            var cells = stale.Where(x => x.Item1 == i).Select(x => x.Item2).ToHashSet();
            if (cells.Count > 0) _sheets[i].ClearCached(cells);
        }

        Recompute(stale, formulas, byName);
        _changed.Clear();
    }

    /// <summary>Work out the given formula cells again, in whatever order their dependencies allow.</summary>
    private void RecomputeThese(HashSet<(int, CellRef)> cells)
    {
        if (cells.Count == 0) return;
        var byName = new Dictionary<string, int>(StringComparer.CurrentCultureIgnoreCase);
        for (int i = 0; i < _sheets.Count; i++) byName[_sheets[i].Name] = i;

        var formulas = new List<(int Sheet, CellRef Ref, RefRange[] Reads)>();
        for (int i = 0; i < _sheets.Count; i++)
            foreach (var (reference, formula) in _sheets[i].Formulas())
                formulas.Add((i, reference, Refs.Parse(formula).ToArray()));

        Recompute(cells, formulas, byName);
    }

    /// <summary>
    /// Work out the totals that just went stale, so the sheet shows numbers rather than the formulas behind them.
    /// A formula is only computed once everything it reads has been, and anything Plain does not fully understand is
    /// left with no value at all: Excel recalculates the file when it opens it, and no number beats a wrong one.
    /// </summary>
    private void Recompute(HashSet<(int, CellRef)> stale,
                           List<(int Sheet, CellRef Ref, RefRange[] Reads)> formulas,
                           Dictionary<string, int> byName)
    {
        if (stale.Count == 0) return;

        var pending = formulas.Where(f => stale.Contains((f.Sheet, f.Ref))).ToList();
        if (pending.Count == 0) return;

        var known = new Dictionary<(int, CellRef), Value>();
        var waiting = new HashSet<(int, CellRef)>(pending.Select(f => (f.Sheet, f.Ref)));

        Value Look(int from, string? sheetName, CellRef cell)
        {
            int sheet = from;
            if (sheetName is not null && !byName.TryGetValue(sheetName, out sheet)) return Value.Blank;
            if (known.TryGetValue((sheet, cell), out var found)) return found;
            return _sheets[sheet].ValueOf(cell);
        }

        // Each pass settles at least one formula or nothing more can be settled, so this cannot run away.
        for (int pass = 0; pass < pending.Count + 1 && waiting.Count > 0; pass++)
        {
            bool progressed = false;
            foreach (var (sheet, reference, reads) in pending.ToList())
            {
                if (!waiting.Contains((sheet, reference))) continue;

                // Wait until nothing it reads is still to be worked out.
                bool ready = true;
                foreach (var range in reads)
                {
                    int on = sheet;
                    if (range.Sheet is not null && !byName.TryGetValue(range.Sheet, out on)) continue;
                    foreach (var still in waiting)
                        if (still.Item1 == on && range.Contains(still.Item2) && still != (sheet, reference)) { ready = false; break; }
                    if (!ready) break;
                }
                if (!ready) continue;

                var formula = _sheets[sheet].Read(reference).Formula;
                waiting.Remove((sheet, reference));
                progressed = true;
                if (formula is null) continue;

                if (Formula.TryEvaluate(formula, (name, cell) => Look(sheet, name, cell), out var value))
                {
                    known[(sheet, reference)] = value;
                    _sheets[sheet].WriteCached(reference, value);
                }
            }
            if (!progressed) break;   // what is left reads itself, or reads something that does
        }
    }

    private static bool ReadsAnything(RefRange[] reads, int ownSheet,
                                      Dictionary<string, int> byName,
                                      Dictionary<int, HashSet<CellRef>> dirty)
    {
        foreach (var range in reads)
        {
            int sheet = ownSheet;
            if (range.Sheet is not null && !byName.TryGetValue(range.Sheet, out sheet)) continue;
            if (!dirty.TryGetValue(sheet, out var cells) || cells.Count == 0) continue;

            long width = (long)range.ColumnMax - range.ColumnMin + 1;
            long height = (long)range.RowMax - range.RowMin + 1;
            if (width * height <= 64)
            {
                // A small range: ask the set about each of its cells.
                for (int c = range.ColumnMin; c <= range.ColumnMax; c++)
                    for (int r = range.RowMin; r <= range.RowMax; r++)
                        if (cells.Contains(new CellRef(c, r))) return true;
            }
            else
            {
                // A big one, a whole column say: walk the handful of changed cells instead.
                foreach (var cell in cells) if (range.Contains(cell)) return true;
            }
        }
        return false;
    }

    /// <summary>
    /// Put a row or column in, or take one out, on one sheet, and make every formula in the workbook still mean what
    /// it meant. A formula that pointed only at what was deleted becomes #REF!, the way Excel does it, because a
    /// visible error is better than a number that quietly went wrong.
    /// </summary>
    public int Apply(Sheet sheet, GridEdit edit, int at)
    {
        if (at < 1) throw new ArgumentOutOfRangeException(nameof(at), "Rows and columns are numbered from one.");
        int limit = edit is GridEdit.InsertRow or GridEdit.DeleteRow ? 1048576 : 16384;
        if (at > limit) throw new ArgumentOutOfRangeException(nameof(at), "That is past the end of the sheet.");

        sheet.ShiftCells(edit, at);
        int adjusted = 0;
        var emptied = new HashSet<(int, CellRef)>();
        for (int i = 0; i < _sheets.Count; i++)
        {
            var here = new List<CellRef>();
            adjusted += _sheets[i].AdjustFormulas(edit, at, sheet.Name, ownSheet: _sheets[i] == sheet, here);
            foreach (var cell in here) emptied.Add((i, cell));
        }

        // A formula that had to move still has to show a number, so work the moved ones out again.
        RecomputeThese(emptied);
        RequestFullRecalculation();
        return adjusted;
    }

    /// <summary>Write every part Plain changed back into the package. Nothing else is touched.</summary>
    public void Flush()
    {
        ClearStaleCaches();
        foreach (var sheet in _sheets) sheet.Flush();
        _styles?.Flush();
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

    private Dialect D => _book.Dialect;

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
            _rowIndex = null;
            _extent = null;
            _doc = Xml.Parse(_book.Package.Read(PartName));
            _data = _doc.Root!.Element(D.Sheet + "sheetData")
                    ?? throw new OpcPackage.PackageException($"The sheet \"{Name}\" has no cell data.");
            return _data;
        }
    }

    private CellRef? _extent;

    /// <summary>The furthest cell that holds anything, which is how wide and tall Plain draws the grid.</summary>
    public CellRef Extent
    {
        get
        {
            if (_extent is { } known) return known;
            int maxCol = 1, maxRow = 1;
            foreach (var row in Data.Elements(D.Sheet + "row"))
                foreach (var c in row.Elements(D.Sheet + "c"))
                    if (CellRef.TryParse((string?)c.Attribute("r") ?? "", out var r) && !string.IsNullOrEmpty(c.Value))
                    {
                        if (r.Column > maxCol) maxCol = r.Column;
                        if (r.Row > maxRow) maxRow = r.Row;
                    }
            _extent = new CellRef(maxCol, maxRow);
            return _extent.Value;
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
        var format = sheetElement?.Element(D.Sheet + "sheetFormatPr");
        if (format is not null && double.TryParse((string?)format.Attribute("defaultColWidth"),
                System.Globalization.NumberStyles.Float, CultureInfo.InvariantCulture, out var d) && d > 0)
            _defaultWidth = d;
        foreach (var c in sheetElement?.Element(D.Sheet + "cols")?.Elements(D.Sheet + "col") ?? Enumerable.Empty<XElement>())
        {
            if (!int.TryParse((string?)c.Attribute("min"), out var min)) continue;
            if (!int.TryParse((string?)c.Attribute("max"), out var max)) continue;
            if (!double.TryParse((string?)c.Attribute("width"), System.Globalization.NumberStyles.Float, CultureInfo.InvariantCulture, out var w)) continue;
            list.Add((min, Math.Min(max, 16384), w));
        }
        return list;
    }

    // Rows are looked up by number rather than scanned for. Without this, drawing one screen of a twenty thousand
    // row sheet walked the whole sheet for every cell, and searching one was quadratic: twenty seconds, not twenty
    // milliseconds.
    private Dictionary<int, XElement>? _rowIndex;

    private Dictionary<int, XElement> RowIndex
    {
        get
        {
            if (_rowIndex is not null) return _rowIndex;
            // Load the sheet before building, because loading it clears this very field.
            var data = Data;
            var index = new Dictionary<int, XElement>();
            foreach (var row in data.Elements(D.Sheet + "row"))
                if (Xml.Int(row.Attribute("r")) is { } number) index[number] = row;
            _rowIndex = index;
            return index;
        }
    }

    private XElement? FindRow(int row) => RowIndex.TryGetValue(row, out var found) ? found : null;

    private XElement? FindCell(CellRef cell)
    {
        string want = cell.ToString();
        return FindRow(cell.Row)?.Elements(D.Sheet + "c").FirstOrDefault(c => (string?)c.Attribute("r") == want);
    }

    public Cell Read(string reference) => Read(CellRef.Parse(reference));

    public Cell Read(CellRef reference)
    {
        var c = FindCell(reference);
        return c is null ? new Cell(reference, CellKind.Empty, "", "", null) : ReadFrom(c, reference);
    }

    private Cell ReadFrom(XElement c, CellRef reference)
    {
        string? type = (string?)c.Attribute("t");
        string? formula = c.Element(D.Sheet + "f")?.Value;
        var v = c.Element(D.Sheet + "v");

        string raw, display;
        CellKind kind;
        switch (type)
        {
            case "s":
                raw = display = int.TryParse(v?.Value, NumberStyles.Integer, CultureInfo.InvariantCulture, out var idx) ? _book.SharedString(idx) : "";
                kind = CellKind.Text;
                break;
            case "inlineStr":
                raw = display = c.Element(D.Sheet + "is")?.Value ?? "";
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
                display = _book.Styles.Format(raw, Xml.Int(c.Attribute("s"), 0));
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

    /// <summary>Every non-empty cell, in reading order. Each one is read from the element in hand, never looked up again.</summary>
    public IEnumerable<Cell> Cells()
    {
        foreach (var row in Data.Elements(D.Sheet + "row"))
            foreach (var c in row.Elements(D.Sheet + "c"))
                if (CellRef.TryParse((string?)c.Attribute("r") ?? "", out var r))
                {
                    var cell = ReadFrom(c, r);
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
        c.Elements(D.Sheet + "f").Remove();
        c.Elements(D.Sheet + "v").Remove();
        c.Elements(D.Sheet + "is").Remove();
        c.Attribute("t")?.Remove();

        if (typed.StartsWith('=') && typed.Length > 1)
        {
            c.Add(new XElement(D.Sheet + "f", typed[1..]));
            _book.RequestFullRecalculation();
        }
        else if (typed.Length == 0)
        {
            // An emptied cell keeps its formatting and loses its content, the way Delete behaves in Excel.
        }
        else if (typed.StartsWith('\''))
        {
            // A leading apostrophe is the old way of saying "this is text, whatever it looks like".
            c.SetAttributeValue("t", "s");
            c.Add(new XElement(D.Sheet + "v", _book.InternString(typed[1..]).ToString(CultureInfo.InvariantCulture)));
            _book.RequestFullRecalculation();
        }
        else if (double.TryParse(typed, NumberStyles.Float, CultureInfo.InvariantCulture, out var number)
                 && !typed.Contains(' ') && double.IsFinite(number))
        {
            c.Add(new XElement(D.Sheet + "v", number.ToString("R", CultureInfo.InvariantCulture)));
            _book.RequestFullRecalculation();
        }
        else
        {
            c.SetAttributeValue("t", "s");
            c.Add(new XElement(D.Sheet + "v", _book.InternString(typed).ToString(CultureInfo.InvariantCulture)));
            _book.RequestFullRecalculation();
        }
        _dirty = true;
        _extent = null;
        _book.NoteChanged(this, reference);
    }

    public void Set(string reference, string typed) => Set(CellRef.Parse(reference), typed);

    /// <summary>
    /// Show the cells the given way: as currency, as a percentage, as a date. The cell keeps everything else about
    /// how it looks. An empty code puts it back to however the format shows a number with no instructions.
    /// </summary>
    public void SetFormat(IEnumerable<CellRef> cells, string code)
    {
        foreach (var reference in cells)
        {
            var cell = EnsureCell(reference);
            int current = Xml.Int(cell.Attribute("s"), 0);
            int style = _book.Styles.WithFormat(current, code);
            cell.SetAttributeValue("s", style);
        }
        _dirty = true;
    }

    /// <summary>Make the cells bold or italic, or take it off, keeping everything else about their font.</summary>
    public void SetWeight(IEnumerable<CellRef> cells, bool? bold, bool? italic)
    {
        foreach (var reference in cells)
        {
            var cell = EnsureCell(reference);
            int current = Xml.Int(cell.Attribute("s"), 0);
            cell.SetAttributeValue("s", _book.Styles.WithWeight(current, bold, italic));
        }
        _dirty = true;
    }

    /// <summary>Is this cell bold, or italic?</summary>
    public bool HasWeight(CellRef reference, string mark)
    {
        var cell = FindCell(reference);
        return cell is not null && _book.Styles.HasWeight(Xml.Int(cell.Attribute("s"), 0), mark);
    }

    /// <summary>How the cell is shown now, as a format code; empty when it has no instructions.</summary>
    public string FormatOf(CellRef reference)
    {
        var cell = FindCell(reference);
        return cell is null ? "" : _book.Styles.CodeAt(Xml.Int(cell.Attribute("s"), 0));
    }

    private XElement EnsureCell(CellRef reference)
    {
        var row = FindRow(reference.Row);
        if (row is null)
        {
            row = new XElement(D.Sheet + "row", new XAttribute("r", reference.Row));
            var after = Data.Elements(D.Sheet + "row").LastOrDefault(r => Xml.Int(r.Attribute("r"), 0) < reference.Row);
            if (after is not null) after.AddAfterSelf(row); else Data.AddFirst(row);
            RowIndex[reference.Row] = row;
        }

        string want = reference.ToString();
        var cell = row.Elements(D.Sheet + "c").FirstOrDefault(c => (string?)c.Attribute("r") == want);
        if (cell is not null) return cell;

        cell = new XElement(D.Sheet + "c", new XAttribute("r", want));
        // Cells must stay in ascending column order or Excel reports the file as damaged.
        var prev = row.Elements(D.Sheet + "c")
            .LastOrDefault(c => CellRef.TryParse((string?)c.Attribute("r") ?? "", out var r) && r.Column < reference.Column);
        if (prev is not null) prev.AddAfterSelf(cell); else row.AddFirst(cell);
        return cell;
    }

    /// <summary>Every formula on this sheet, with the cell it sits in. Used to work out what an edit made untrue.</summary>
    internal IReadOnlyList<(CellRef Ref, string Formula)> Formulas()
    {
        var list = new List<(CellRef, string)>();
        foreach (var row in Data.Elements(D.Sheet + "row"))
            foreach (var c in row.Elements(D.Sheet + "c"))
            {
                var f = c.Element(D.Sheet + "f");
                if (f is null) continue;
                if (CellRef.TryParse((string?)c.Attribute("r") ?? "", out var reference)) list.Add((reference, f.Value));
            }
        return list;
    }

    /// <summary>
    /// Drop the value Excel last cached beside the given formulas. Plain does not evaluate formulas, so a cached total
    /// that depends on an edited cell is a number that may no longer be true; with no cached value, whatever opens the
    /// file has to work it out. Formulas nothing touched keep their values, and so does every cell without a formula.
    /// </summary>
    internal bool ClearCached(HashSet<CellRef> cells)
    {
        if (cells.Count == 0) return false;
        bool any = false;
        foreach (var row in Data.Elements(D.Sheet + "row"))
            foreach (var c in row.Elements(D.Sheet + "c"))
            {
                if (c.Element(D.Sheet + "f") is null) continue;
                if (!CellRef.TryParse((string?)c.Attribute("r") ?? "", out var reference) || !cells.Contains(reference)) continue;
                if (c.Element(D.Sheet + "v") is null) continue;
                c.Elements(D.Sheet + "v").Remove();
                // A cached string result is typed on the cell; without the value the type means nothing.
                if ((string?)c.Attribute("t") == "str") c.Attribute("t")!.Remove();
                any = true;
            }
        if (any) _dirty = true;
        return any;
    }

    /// <summary>
    /// Move the cells for an inserted or deleted row or column. Rows and cells are renamed rather than rebuilt, so a
    /// cell keeps its style, its formula and everything else about it; only where it sits changes.
    /// </summary>
    internal void ShiftCells(GridEdit edit, int at)
    {
        var data = Data;

        if (edit == GridEdit.DeleteRow)
            foreach (var row in data.Elements(D.Sheet + "row").Where(r => Xml.Int(r.Attribute("r"), 0) == at).ToList())
                row.Remove();

        foreach (var row in data.Elements(D.Sheet + "row").ToList())
        {
            int number = Xml.Int(row.Attribute("r"), 0);
            if (number == 0) continue;

            if (edit == GridEdit.DeleteColumn)
                foreach (var cell in row.Elements(D.Sheet + "c")
                             .Where(c => CellRef.TryParse((string?)c.Attribute("r") ?? "", out var r) && r.Column == at).ToList())
                    cell.Remove();

            int moved = edit switch
            {
                GridEdit.InsertRow => number >= at ? number + 1 : number,
                GridEdit.DeleteRow => number > at ? number - 1 : number,
                _ => number,
            };
            if (moved != number) row.SetAttributeValue("r", moved);

            foreach (var cell in row.Elements(D.Sheet + "c"))
            {
                if (!CellRef.TryParse((string?)cell.Attribute("r") ?? "", out var reference)) continue;
                var landed = Grid.Move(reference, edit, at);
                if (landed is { } to && to != reference) cell.SetAttributeValue("r", to.ToString());
            }

            // The spans hint and the sheet dimension describe a shape that has just changed; Excel works both out
            // again, and a stale one is worse than none.
            row.Attribute("spans")?.Remove();
        }

        if (edit is GridEdit.InsertColumn or GridEdit.DeleteColumn) ShiftColumnWidths(edit, at);
        _doc?.Root?.Element(D.Sheet + "dimension")?.Remove();

        _rowIndex = null;
        _extent = null;
        _widths = null;
        _dirty = true;
    }

    private void ShiftColumnWidths(GridEdit edit, int at)
    {
        var cols = _doc?.Root?.Element(D.Sheet + "cols");
        if (cols is null) return;
        foreach (var col in cols.Elements(D.Sheet + "col").ToList())
        {
            int min = Xml.Int(col.Attribute("min"), 0), max = Xml.Int(col.Attribute("max"), 0);
            if (min == 0 || max == 0) continue;
            if (edit == GridEdit.InsertColumn)
            {
                if (min >= at) col.SetAttributeValue("min", min + 1);
                if (max >= at) col.SetAttributeValue("max", Math.Min(16384, max + 1));
            }
            else
            {
                if (min > at) col.SetAttributeValue("min", min - 1);
                if (max >= at) col.SetAttributeValue("max", max - 1);
                if (Xml.Int(col.Attribute("max"), 0) < Xml.Int(col.Attribute("min"), 1)) col.Remove();
            }
        }
        if (!cols.Elements().Any()) cols.Remove();
    }

    /// <summary>Rewrite every formula on this sheet so it still means what it meant after a row or column moved.</summary>
    internal int AdjustFormulas(GridEdit edit, int at, string targetSheet, bool ownSheet, List<CellRef>? emptied = null)
    {
        int changed = 0;
        foreach (var row in Data.Elements(D.Sheet + "row"))
            foreach (var cell in row.Elements(D.Sheet + "c"))
            {
                var f = cell.Element(D.Sheet + "f");
                if (f is null || f.Value.Length == 0) continue;
                var adjusted = Grid.Adjust(f.Value, edit, at, targetSheet, ownSheet);
                if (adjusted == f.Value) continue;
                f.Value = adjusted;
                cell.Elements(D.Sheet + "v").Remove();   // whatever it worked out to is no longer what it means
                if ((string?)cell.Attribute("t") == "str") cell.Attribute("t")!.Remove();
                if (CellRef.TryParse((string?)cell.Attribute("r") ?? "", out var where)) emptied?.Add(where);
                changed++;
            }
        if (changed > 0) _dirty = true;
        return changed;
    }

    /// <summary>Put a worked-out value beside a formula, the way Excel stores the answer it last calculated.</summary>
    internal void WriteCached(CellRef reference, Value value)
    {
        var cell = FindCell(reference);
        if (cell is null) return;
        cell.Elements(D.Sheet + "v").Remove();
        cell.Attribute("t")?.Remove();

        switch (value.Kind)
        {
            case Value.Sort.Number:
                if (!double.IsFinite(value.Number)) return;   // nothing sensible to store
                cell.Add(new XElement(D.Sheet + "v", value.Number.ToString("R", CultureInfo.InvariantCulture)));
                break;
            case Value.Sort.Bool:
                cell.SetAttributeValue("t", "b");
                cell.Add(new XElement(D.Sheet + "v", value.Number != 0 ? "1" : "0"));
                break;
            case Value.Sort.Error:
                cell.SetAttributeValue("t", "e");
                cell.Add(new XElement(D.Sheet + "v", value.Text));
                break;
            case Value.Sort.Text:
                cell.SetAttributeValue("t", "str");
                cell.Add(new XElement(D.Sheet + "v", value.Text));
                break;
            default:
                return;   // blank: leave it with no value, which is what "not worked out" looks like
        }
        _dirty = true;
    }

    /// <summary>What this cell stands for when a formula reads it.</summary>
    internal Value ValueOf(CellRef reference)
    {
        var cell = Read(reference);
        return cell.Kind switch
        {
            CellKind.Empty => Value.Blank,
            CellKind.Boolean => Value.Of(cell.Raw == "1"),
            CellKind.Error => Value.Error(cell.Raw),
            CellKind.Text => Value.Of(cell.Display),
            _ => double.TryParse(cell.Raw, NumberStyles.Float, CultureInfo.InvariantCulture, out var n)
                 ? Value.Of(n)
                 : cell.Raw.Length == 0 ? Value.Blank : Value.Of(cell.Raw),
        };
    }

    internal void Flush()
    {
        if (!_dirty || _doc is null) return;
        _book.Package.Write(PartName, Xml.ToBytes(_doc));
        _dirty = false;
    }
}
