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

    /// <summary>
    /// The style index that is the one a cell has, but bold or not, italic or not. The font is copied and altered
    /// rather than replaced, so a cell in a particular typeface or colour keeps it.
    /// </summary>
    public int WithWeight(int currentStyle, bool? bold, bool? italic)
    {
        if (_doc?.Root is null) throw new OpcPackage.PackageException("This workbook has no style table, so Plain cannot change how a cell looks.");
        var cellXfs = _doc.Root.Element(D.Sheet + "cellXfs") ?? throw new OpcPackage.PackageException("This workbook's style table has no cell formats.");
        var fonts = _doc.Root.Element(D.Sheet + "fonts") ?? throw new OpcPackage.PackageException("This workbook's style table has no fonts.");

        var all = cellXfs.Elements(D.Sheet + "xf").ToList();
        var basis = currentStyle >= 0 && currentStyle < all.Count ? all[currentStyle] : all.FirstOrDefault();
        if (basis is null) throw new OpcPackage.PackageException("This workbook's style table is empty.");

        int fontIndex = Xml.Int(basis.Attribute("fontId"), 0);
        var allFonts = fonts.Elements(D.Sheet + "font").ToList();
        var font = fontIndex >= 0 && fontIndex < allFonts.Count ? allFonts[fontIndex] : allFonts.FirstOrDefault();
        if (font is null) throw new OpcPackage.PackageException("This workbook's style table has no fonts.");

        var wanted = new XElement(font);
        Mark(wanted, "b", bold);
        Mark(wanted, "i", italic);

        int newFont = allFonts.FindIndex(f => Same(f, wanted));
        if (newFont < 0)
        {
            fonts.Add(wanted);
            fonts.SetAttributeValue("count", allFonts.Count + 1);
            newFont = allFonts.Count;
            _dirty = true;
        }

        for (int i = 0; i < all.Count; i++)
            if (Xml.Int(all[i].Attribute("fontId"), 0) == newFont
                && Xml.Int(all[i].Attribute("numFmtId"), 0) == Xml.Int(basis.Attribute("numFmtId"), 0)
                && (string?)all[i].Attribute("fillId") == (string?)basis.Attribute("fillId")
                && (string?)all[i].Attribute("borderId") == (string?)basis.Attribute("borderId"))
                return i;

        var made = new XElement(basis);
        made.SetAttributeValue("fontId", newFont);
        made.SetAttributeValue("applyFont", "1");
        cellXfs.Add(made);
        cellXfs.SetAttributeValue("count", all.Count + 1);
        if (_formats.TryGetValue(currentStyle, out var carried)) _formats[all.Count] = carried;
        _dirty = true;
        return all.Count;

        // A font's children have to come in the order the format lays down - bold, then italic, then everything else -
        // and Excel refuses a file that puts them anywhere else. So both are lifted out and put back at the front.
        void Mark(XElement target, string name, bool? on)
        {
            if (on is null) return;
            bool bold = target.Element(D.Sheet + "b") is not null;
            bool italic = target.Element(D.Sheet + "i") is not null;
            if (name == "b") bold = on.Value; else italic = on.Value;

            target.Elements(D.Sheet + "b").Remove();
            target.Elements(D.Sheet + "i").Remove();
            if (italic) target.AddFirst(new XElement(D.Sheet + "i"));
            if (bold) target.AddFirst(new XElement(D.Sheet + "b"));
        }

        static bool Same(XElement a, XElement b) => XNode.DeepEquals(a, b);
    }

    /// <summary>
    /// How a cell's contents sit in it: left, centre or right, and whether long text wraps onto more lines rather
    /// than running under the next cell. Alignment lives on the cell format itself, not on the font, so this is a
    /// smaller change than bold: find or make a format that is this one plus the alignment asked for.
    /// </summary>
    public int WithAlignment(int currentStyle, string? horizontal, bool? wrap)
    {
        var (cellXfs, all, basis) = Formats(currentStyle);

        var wanted = new XElement(basis);
        var alignment = wanted.Element(D.Sheet + "alignment");
        if (alignment is null)
        {
            alignment = new XElement(D.Sheet + "alignment");
            // alignment comes after the attributes and before protection, and Excel minds the order.
            var protection = wanted.Element(D.Sheet + "protection");
            if (protection is not null) protection.AddBeforeSelf(alignment); else wanted.Add(alignment);
        }

        if (horizontal is not null)
        {
            // "general" is what a cell has when nobody has chosen, and it is said by saying nothing.
            if (horizontal is "general" or "") alignment.SetAttributeValue("horizontal", null);
            else alignment.SetAttributeValue("horizontal", horizontal);
        }
        if (wrap is not null) alignment.SetAttributeValue("wrapText", wrap.Value ? "1" : null);

        if (!alignment.HasAttributes) alignment.Remove();
        wanted.SetAttributeValue("applyAlignment", "1");

        return Adopt(cellXfs, all, wanted, currentStyle);
    }

    /// <summary>What the cell says about where its contents sit, for showing the state of a button.</summary>
    public (string Horizontal, bool Wrap) AlignmentAt(int styleIndex)
    {
        var all = _doc?.Root?.Element(D.Sheet + "cellXfs")?.Elements(D.Sheet + "xf").ToList();
        if (all is null || styleIndex < 0 || styleIndex >= all.Count) return ("general", false);
        var alignment = all[styleIndex].Element(D.Sheet + "alignment");
        if (alignment is null) return ("general", false);
        return ((string?)alignment.Attribute("horizontal") ?? "general",
                (string?)alignment.Attribute("wrapText") is "1" or "true");
    }

    /// <summary>
    /// The colour behind a cell, and the colour of its words, as six hex digits. Null for either leaves it alone,
    /// and an empty string takes it back to none. A fill in Excel is a pattern with a foreground colour, and the
    /// first two fills in every workbook are reserved, which is why a new one is always added rather than reused
    /// from the front of the list.
    /// </summary>
    public int WithColours(int currentStyle, string? background, string? ink)
    {
        var (cellXfs, all, basis) = Formats(currentStyle);
        var wanted = new XElement(basis);

        if (background is not null)
        {
            var fills = _doc!.Root!.Element(D.Sheet + "fills")
                ?? throw new OpcPackage.PackageException("This workbook's style table has no fills.");
            var allFills = fills.Elements(D.Sheet + "fill").ToList();

            XElement made;
            if (background.Length == 0)
            {
                made = new XElement(D.Sheet + "fill", new XElement(D.Sheet + "patternFill",
                    new XAttribute("patternType", "none")));
            }
            else
            {
                made = new XElement(D.Sheet + "fill", new XElement(D.Sheet + "patternFill",
                    new XAttribute("patternType", "solid"),
                    new XElement(D.Sheet + "fgColor", new XAttribute("rgb", "FF" + Hex(background))),
                    new XElement(D.Sheet + "bgColor", new XAttribute("indexed", "64"))));
            }

            int index = allFills.FindIndex(f => XNode.DeepEquals(f, made));
            if (index < 0)
            {
                fills.Add(made);
                fills.SetAttributeValue("count", allFills.Count + 1);
                index = allFills.Count;
                _dirty = true;
            }
            wanted.SetAttributeValue("fillId", index);
            wanted.SetAttributeValue("applyFill", "1");
        }

        if (ink is not null)
        {
            var fonts = _doc!.Root!.Element(D.Sheet + "fonts")
                ?? throw new OpcPackage.PackageException("This workbook's style table has no fonts.");
            var allFonts = fonts.Elements(D.Sheet + "font").ToList();
            int fontIndex = Xml.Int(basis.Attribute("fontId"), 0);
            var font = fontIndex >= 0 && fontIndex < allFonts.Count ? allFonts[fontIndex] : allFonts.FirstOrDefault();
            if (font is null) throw new OpcPackage.PackageException("This workbook's style table has no fonts.");

            var newFont = new XElement(font);
            newFont.Elements(D.Sheet + "color").Remove();
            if (ink.Length > 0)
            {
                // A font's colour goes after b and i and before sz, name and the rest.
                var colour = new XElement(D.Sheet + "color", new XAttribute("rgb", "FF" + Hex(ink)));
                var after = newFont.Elements(D.Sheet + "i").LastOrDefault()
                         ?? newFont.Elements(D.Sheet + "b").LastOrDefault();
                if (after is not null) after.AddAfterSelf(colour); else newFont.AddFirst(colour);
            }

            int index = allFonts.FindIndex(f => XNode.DeepEquals(f, newFont));
            if (index < 0)
            {
                fonts.Add(newFont);
                fonts.SetAttributeValue("count", allFonts.Count + 1);
                index = allFonts.Count;
                _dirty = true;
            }
            wanted.SetAttributeValue("fontId", index);
            wanted.SetAttributeValue("applyFont", "1");
        }

        return Adopt(cellXfs, all, wanted, currentStyle);
    }

    /// <summary>
    /// Lines round a cell. Which sides, how heavy, and what colour. An empty style takes the lines away.
    ///
    /// A border in the file is a set of five sides in a fixed order, and Excel refuses a file that puts them in
    /// any other, so this always rebuilds all five rather than editing one in place.
    /// </summary>
    public int WithBorder(int currentStyle, IReadOnlyCollection<string> sides, string style, string colour)
    {
        var (cellXfs, all, basis) = Formats(currentStyle);
        var borders = _doc!.Root!.Element(D.Sheet + "borders")
            ?? throw new OpcPackage.PackageException("This workbook's style table has no borders.");
        var allBorders = borders.Elements(D.Sheet + "border").ToList();

        int borderId = Xml.Int(basis.Attribute("borderId"), 0);
        var from = borderId >= 0 && borderId < allBorders.Count ? allBorders[borderId] : allBorders.FirstOrDefault();
        var wanted = from is null ? new XElement(D.Sheet + "border") : new XElement(from);

        // Every side, in the order the format lays down, whether or not this call touches it.
        var keep = new Dictionary<string, XElement?>();
        foreach (var side in Order) keep[side] = wanted.Element(D.Sheet + side);
        wanted.RemoveNodes();

        foreach (var side in Order)
        {
            var element = new XElement(D.Sheet + side);
            bool touching = sides.Contains(side, StringComparer.OrdinalIgnoreCase)
                         || (sides.Contains("all", StringComparer.OrdinalIgnoreCase) && side != "diagonal");

            if (touching && style.Length > 0)
            {
                element.SetAttributeValue("style", style);
                if (colour.Length > 0)
                    element.Add(new XElement(D.Sheet + "color", new XAttribute("rgb", "FF" + Hex(colour))));
            }
            else if (!touching && keep[side] is { } was)
            {
                // A side this call says nothing about keeps whatever it had.
                element = new XElement(was);
            }
            wanted.Add(element);
        }

        int index = allBorders.FindIndex(b => XNode.DeepEquals(b, wanted));
        if (index < 0)
        {
            borders.Add(wanted);
            borders.SetAttributeValue("count", allBorders.Count + 1);
            index = allBorders.Count;
            _dirty = true;
        }

        var made = new XElement(basis);
        made.SetAttributeValue("borderId", index);
        made.SetAttributeValue("applyBorder", "1");
        return Adopt(cellXfs, all, made, currentStyle);
    }

    /// <summary>The order the five sides of a border have to be written in.</summary>
    private static readonly string[] Order = { "left", "right", "top", "bottom", "diagonal" };

    /// <summary>Which sides of this cell have a line, and how heavy each is.</summary>
    public IReadOnlyList<(string Side, string Style)> BorderAt(int styleIndex)
    {
        var found = new List<(string, string)>();
        var root = _doc?.Root;
        var all = root?.Element(D.Sheet + "cellXfs")?.Elements(D.Sheet + "xf").ToList();
        if (root is null || all is null || styleIndex < 0 || styleIndex >= all.Count) return found;

        var borders = root.Element(D.Sheet + "borders")?.Elements(D.Sheet + "border").ToList();
        int id = Xml.Int(all[styleIndex].Attribute("borderId"), 0);
        if (borders is null || id < 0 || id >= borders.Count) return found;

        foreach (var side in Order)
        {
            var style = (string?)borders[id].Element(D.Sheet + side)?.Attribute("style");
            if (!string.IsNullOrEmpty(style)) found.Add((side, style));
        }
        return found;
    }

    /// <summary>The weights people ask for, and what the file calls them.</summary>
    public static readonly (string Name, string Code)[] Weights =
    {
        ("Thin", "thin"),
        ("Medium", "medium"),
        ("Thick", "thick"),
        ("Dotted", "dotted"),
        ("Dashed", "dashed"),
        ("Double", "double"),
    };

    /// <summary>The colours a cell is wearing, as six hex digits, or empty where it wears none.</summary>
    public (string Background, string Ink) ColoursAt(int styleIndex)
    {
        var root = _doc?.Root;
        var all = root?.Element(D.Sheet + "cellXfs")?.Elements(D.Sheet + "xf").ToList();
        if (root is null || all is null || styleIndex < 0 || styleIndex >= all.Count) return ("", "");
        var xf = all[styleIndex];

        string background = "";
        var fills = root.Element(D.Sheet + "fills")?.Elements(D.Sheet + "fill").ToList();
        int fillId = Xml.Int(xf.Attribute("fillId"), 0);
        if (fills is not null && fillId >= 0 && fillId < fills.Count)
        {
            var pattern = fills[fillId].Element(D.Sheet + "patternFill");
            if ((string?)pattern?.Attribute("patternType") == "solid")
                background = Six((string?)pattern.Element(D.Sheet + "fgColor")?.Attribute("rgb"));
        }

        string ink = "";
        var fonts = root.Element(D.Sheet + "fonts")?.Elements(D.Sheet + "font").ToList();
        int fontId = Xml.Int(xf.Attribute("fontId"), 0);
        if (fonts is not null && fontId >= 0 && fontId < fonts.Count)
            ink = Six((string?)fonts[fontId].Element(D.Sheet + "color")?.Attribute("rgb"));

        return (background, ink);
    }

    /// <summary>Six hex digits, whatever was written: with or without a leading alpha pair, with or without a hash.</summary>
    private static string Hex(string colour)
    {
        var clean = new string(colour.Where(Uri.IsHexDigit).ToArray()).ToUpperInvariant();
        if (clean.Length == 8) clean = clean[2..];
        return clean.Length == 6 ? clean : "000000";
    }

    private static string Six(string? rgb) =>
        rgb is null ? "" : rgb.Length == 8 ? rgb[2..].ToUpperInvariant() : rgb.ToUpperInvariant();

    /// <summary>The cell formats, with the one this cell is using, or a clear word about why there are none.</summary>
    private (XElement CellXfs, List<XElement> All, XElement Basis) Formats(int currentStyle)
    {
        if (_doc?.Root is null) throw new OpcPackage.PackageException("This workbook has no style table, so Plain cannot change how a cell looks.");
        var cellXfs = _doc.Root.Element(D.Sheet + "cellXfs") ?? throw new OpcPackage.PackageException("This workbook's style table has no cell formats.");
        var all = cellXfs.Elements(D.Sheet + "xf").ToList();
        // Build on what the cell is already wearing, so setting a colour does not throw away its number format.
        var basis = currentStyle >= 0 && currentStyle < all.Count ? all[currentStyle] : all.FirstOrDefault();
        if (basis is null) throw new OpcPackage.PackageException("This workbook's style table is empty.");
        return (cellXfs, all, basis);
    }

    /// <summary>Use a format that already says this, or add it. Two identical formats would only bloat the file.</summary>
    private int Adopt(XElement cellXfs, List<XElement> all, XElement wanted, int from)
    {
        for (int i = 0; i < all.Count; i++)
            if (XNode.DeepEquals(all[i], wanted)) return i;

        cellXfs.Add(wanted);
        cellXfs.SetAttributeValue("count", all.Count + 1);
        if (_formats.TryGetValue(from, out var carried)) _formats[all.Count] = carried;
        _dirty = true;
        return all.Count;
    }

    /// <summary>Is the cell's font bold, or italic?</summary>
    public bool HasWeight(int styleIndex, string mark)
    {
        var cellXfs = _doc?.Root?.Element(D.Sheet + "cellXfs");
        var fonts = _doc?.Root?.Element(D.Sheet + "fonts");
        if (cellXfs is null || fonts is null) return false;
        var all = cellXfs.Elements(D.Sheet + "xf").ToList();
        if (styleIndex < 0 || styleIndex >= all.Count) return false;
        int fontIndex = Xml.Int(all[styleIndex].Attribute("fontId"), 0);
        var allFonts = fonts.Elements(D.Sheet + "font").ToList();
        return fontIndex >= 0 && fontIndex < allFonts.Count && allFonts[fontIndex].Element(D.Sheet + mark) is not null;
    }

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
