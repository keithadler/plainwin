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

    private readonly Dialect D;

    public Styles(OpcPackage pkg, Dialect dialect)
    {
        D = dialect;
        if (!pkg.Has("xl/styles.xml")) return;
        XDocument doc;
        try { doc = Xml.Parse(pkg.Read("xl/styles.xml")); } catch { return; }

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
