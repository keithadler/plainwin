using System.Globalization;
using System.Xml.Linq;

namespace Plain.Core;

/// <summary>
/// Just enough of xl/styles.xml to show a number the way Excel shows it. Plain covers the built-in formats and the
/// common shapes of custom ones; anything it does not recognise is shown as the plain number rather than guessed at,
/// and the format string itself is never altered.
/// </summary>
public sealed class Styles
{
    private readonly Dictionary<int, string> _formats = new();   // style index -> format code
    private static readonly DateTime Epoch = new(1899, 12, 30);

    private readonly OpcPackage _pkg;
    private XDocument? _doc;
    private bool _dirty;

    /// <summary>The formats worth offering by name, in the order a person would look for them.</summary>
    public static readonly (string Name, string Code)[] Common =
    {
        ("General", ""),
        ("Number", "#,##0.00"),
        ("Whole number", "#,##0"),
        ("Currency", "\u00a3#,##0.00"),
        ("Percent", "0.0%"),
        ("Date", "dd/mm/yyyy"),
        ("Text", "@"),
    };

    private readonly Dialect D;

    public Styles(OpcPackage pkg, Dialect dialect)
    {
        D = dialect;
        _pkg = pkg;
        if (!pkg.Has("xl/styles.xml")) return;
        XDocument doc;
        try { doc = Xml.Parse(pkg.Read("xl/styles.xml")); } catch { return; }
        _doc = doc;

        var custom = new Dictionary<int, string>();
        foreach (var f in doc.Root!.Element(D.Sheet + "numFmts")?.Elements(D.Sheet + "numFmt") ?? Enumerable.Empty<XElement>())
            if (int.TryParse((string?)f.Attribute("numFmtId"), out var id))
                custom[id] = (string?)f.Attribute("formatCode") ?? "";

        int styleIndex = 0;
        foreach (var xf in doc.Root.Element(D.Sheet + "cellXfs")?.Elements(D.Sheet + "xf") ?? Enumerable.Empty<XElement>())
        {
            int id = Xml.Int(xf.Attribute("numFmtId"), 0);
            string code = custom.TryGetValue(id, out var c) ? c : BuiltIn(id);
            if (code.Length > 0) _formats[styleIndex] = code;
            styleIndex++;
        }
    }

    private static string BuiltIn(int id) => id switch
    {
        1 => "0",
        2 => "0.00",
        3 => "#,##0",
        4 => "#,##0.00",
        9 => "0%",
        10 => "0.00%",
        11 => "0.00E+00",
        14 => "d/m/yyyy",
        15 => "d-mmm-yy",
        16 => "d-mmm",
        17 => "mmm-yy",
        18 => "h:mm AM/PM",
        19 => "h:mm:ss AM/PM",
        20 => "h:mm",
        21 => "h:mm:ss",
        22 => "d/m/yyyy h:mm",
        37 or 38 => "#,##0",
        39 or 40 => "#,##0.00",
        44 or 43 => "#,##0.00",
        45 => "mm:ss",
        46 => "[h]:mm:ss",
        47 => "mm:ss.0",
        48 => "##0.0E+0",
        49 => "@",
        _ => "",
    };

    /// <summary>The format code a cell's style asks for, empty when it has none.</summary>
    public string CodeAt(int styleIndex) => _formats.TryGetValue(styleIndex, out var code) ? code : "";

    /// <summary>
    /// The style index that is the one a cell already has, but showing numbers the given way. Everything else about
    /// the cell's look - its font, its colour, its borders - is carried across, because someone changing a column to
    /// currency did not ask to lose the shading on it.
    /// </summary>
    public int WithFormat(int currentStyle, string code)
    {
        if (_doc?.Root is null) throw new OpcPackage.PackageException("This workbook has no style table, so Plain cannot change how a cell looks.");

        int formatId = code.Length == 0 ? 0 : FormatId(code);
        var cellXfs = _doc.Root.Element(D.Sheet + "cellXfs");
        if (cellXfs is null) throw new OpcPackage.PackageException("This workbook's style table has no cell formats.");

        var all = cellXfs.Elements(D.Sheet + "xf").ToList();
        var basis = currentStyle >= 0 && currentStyle < all.Count ? all[currentStyle] : all.FirstOrDefault();
        if (basis is null) throw new OpcPackage.PackageException("This workbook's style table is empty.");

        // If one already says exactly this, use it rather than growing the table every time somebody clicks.
        for (int i = 0; i < all.Count; i++)
            if (SameApartFromFormat(all[i], basis) && Xml.Int(all[i].Attribute("numFmtId"), 0) == formatId)
                return i;

        var made = new XElement(basis);
        made.SetAttributeValue("numFmtId", formatId);
        made.SetAttributeValue("applyNumberFormat", "1");
        cellXfs.Add(made);
        cellXfs.SetAttributeValue("count", all.Count + 1);
        _formats[all.Count] = code;
        _dirty = true;
        return all.Count;
    }

    private static bool SameApartFromFormat(XElement a, XElement b)
    {
        foreach (var name in new[] { "fontId", "fillId", "borderId", "xfId", "applyFont", "applyFill", "applyBorder", "applyAlignment" })
            if ((string?)a.Attribute(name) != (string?)b.Attribute(name)) return false;
        return a.Element(a.Name.Namespace + "alignment")?.ToString() == b.Element(b.Name.Namespace + "alignment")?.ToString();
    }

    /// <summary>The id for a format code, adding it to the table when the workbook has never used it.</summary>
    private int FormatId(string code)
    {
        var numFmts = _doc!.Root!.Element(D.Sheet + "numFmts");
        foreach (var f in numFmts?.Elements(D.Sheet + "numFmt") ?? Enumerable.Empty<XElement>())
            if ((string?)f.Attribute("formatCode") == code) return Xml.Int(f.Attribute("numFmtId"), 0);

        // The built-in codes have fixed ids; using one avoids inventing a format the file already knows.
        for (int id = 0; id <= 49; id++) if (BuiltIn(id) == code) return id;

        if (numFmts is null)
        {
            numFmts = new XElement(D.Sheet + "numFmts", new XAttribute("count", 0));
            _doc.Root.AddFirst(numFmts);
        }
        // Ids below 164 belong to the format itself, so a new one starts above them.
        int next = 164;
        foreach (var f in numFmts.Elements(D.Sheet + "numFmt"))
            next = Math.Max(next, Xml.Int(f.Attribute("numFmtId"), 163) + 1);
        numFmts.Add(new XElement(D.Sheet + "numFmt", new XAttribute("numFmtId", next), new XAttribute("formatCode", code)));
        numFmts.SetAttributeValue("count", numFmts.Elements(D.Sheet + "numFmt").Count());
        _dirty = true;
        return next;
    }

    public void Flush()
    {
        if (!_dirty || _doc is null) return;
        _pkg.Write("xl/styles.xml", Xml.ToBytes(_doc));
        _dirty = false;
    }

    /// <summary>Render a stored number the way its format asks. Text and unrecognised formats come back unchanged.</summary>
    public string Format(string raw, int styleIndex)
    {
        if (!double.TryParse(raw, NumberStyles.Float, CultureInfo.InvariantCulture, out var value)) return raw;
        if (!_formats.TryGetValue(styleIndex, out var code) || code.Length == 0) return General(value);

        // A format can carry up to four sections; the first covers positives and is the one Plain applies.
        string section = code.Split(';')[0].Trim();
        if (section == "@") return raw;
        if (section.Length == 0) return General(value);

        try
        {
            if (LooksLikeDate(section)) return FormatDate(value, section);

            bool percent = section.Contains('%');
            double shown = percent ? value * 100 : value;

            int decimals = Decimals(section);
            bool thousands = section.Contains("#,##") || section.Contains("0,0");
            string number = shown.ToString((thousands ? "N" : "F") + decimals.ToString(CultureInfo.InvariantCulture), CultureInfo.CurrentCulture);

            string prefix = Currency(section);
            return prefix + number + (percent ? "%" : "");
        }
        catch { return General(value); }
    }

    private static string General(double value) =>
        value == Math.Floor(value) && Math.Abs(value) < 1e15
            ? ((long)value).ToString(CultureInfo.CurrentCulture)
            : value.ToString("0.###########", CultureInfo.CurrentCulture);

    private static bool LooksLikeDate(string code)
    {
        bool inLiteral = false;
        foreach (char ch in code)
        {
            if (ch == '"') { inLiteral = !inLiteral; continue; }
            if (inLiteral) continue;
            if (ch is 'y' or 'd' or 'h' or 's') return true;
            if (ch is 'm' or 'M') return true;
        }
        return false;
    }

    private static string FormatDate(double serial, string code)
    {
        var when = Epoch.AddDays(serial);
        // Excel's tokens are close enough to .NET's that a small substitution covers the built-in formats.
        string net = code.Replace("AM/PM", "tt").Replace("am/pm", "tt")
                         .Replace("[h]", "HH").Replace("\\", "");
        net = net.Contains("tt") ? net.Replace("h", "h") : net.Replace("h", "H");
        net = net.Replace("mmmmm", "MMM").Replace("mmmm", "MMMM").Replace("mmm", "MMM");
        // Minutes only follow an hour or precede seconds; elsewhere "m" means month.
        net = FixMinutes(net);
        return when.ToString(net, CultureInfo.CurrentCulture);
    }

    private static string FixMinutes(string code)
    {
        var chars = code.ToCharArray();
        for (int i = 0; i < chars.Length; i++)
        {
            if (chars[i] != 'm') continue;
            int j = i; while (j < chars.Length && chars[j] == 'm') j++;
            bool afterHour = Look(chars, i - 1, "hH");
            bool beforeSecond = Look(chars, j, "s");
            if (!afterHour && !beforeSecond) for (int k = i; k < j; k++) chars[k] = 'M';
            i = j - 1;
        }
        return new string(chars);
    }

    private static bool Look(char[] chars, int from, string wanted)
    {
        int step = from < 0 ? 1 : (from == 0 ? 1 : -1);
        if (from < 0) return false;
        for (int i = from; i >= 0 && i < chars.Length; i += step)
        {
            if (wanted.Contains(chars[i])) return true;
            if (chars[i] is ':' or ' ') continue;
            return false;
        }
        return false;
    }

    private static int Decimals(string code)
    {
        int dot = code.IndexOf('.');
        if (dot < 0) return 0;
        int n = 0;
        for (int i = dot + 1; i < code.Length && (code[i] == '0' || code[i] == '#'); i++) n++;
        return Math.Min(n, 15);
    }

    private static string Currency(string code)
    {
        foreach (char symbol in new[] { '$', '£', '€', '¥' })
            if (code.Contains(symbol)) return symbol.ToString();
        return "";
    }
}
