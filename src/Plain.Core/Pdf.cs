using System.Globalization;
using System.Text;

namespace Plain.Core;

/// <summary>
/// Writes a PDF. More people asked for this than for anything else, because what leaves an office leaves as a PDF:
/// the invoice to the accountant, the worksheet to the cover teacher, the report to the donor.
///
/// It is written out here rather than pulled in, for the same reason as everything else in Plain: no dependency, no
/// installer, and it works the same on any machine. It uses the fonts every PDF reader already has, so nothing is
/// embedded and the file stays small. That does mean the text it can write is the Latin alphabet and the punctuation
/// around it; anything else is refused rather than silently turned into rubbish.
/// </summary>
public sealed class Pdf
{
    public const double A4Width = 595.28, A4Height = 841.89;

    private readonly List<byte[]> _pages = new();
    private readonly MemoryStream _current = new();
    private readonly double _width, _height, _margin;
    private double _y;
    private int _pageCount;

    public Pdf(double width = A4Width, double height = A4Height, double margin = 56)
    {
        _width = width; _height = height; _margin = margin;
        _y = height - margin;
    }

    public double TextWidth => _width - 2 * _margin;

    /// <summary>Is every character one of the fonts can write? Anything else and Plain says so rather than mangling it.</summary>
    public static bool CanWrite(string text) => text.All(c => c is '\t' or '\n' or '\r' || (c >= 32 && c <= 126) || (c >= 160 && c <= 255));

    /// <summary>The characters this cannot write, so the caller can name them.</summary>
    public static string Unwritable(string text) =>
        new(text.Where(c => !(c is '\t' or '\n' or '\r' || (c >= 32 && c <= 126) || (c >= 160 && c <= 255))).Distinct().ToArray());

    // ---------- laying text out ----------

    public void Space(double points) { Room(points); _y -= points; }

    /// <summary>Put a line of text down, wrapping it, and say how far down the page it reached.</summary>
    public void Paragraph(string text, double size = 10.5, bool bold = false, double indent = 0, double after = 4)
    {
        foreach (var line in Wrap(text, size, bold, TextWidth - indent))
        {
            Room(size * 1.35);
            Line(line, _margin + indent, _y - size, size, bold);
            _y -= size * 1.35;
        }
        _y -= after;
    }

    /// <summary>A row of cells at fixed positions, for a table or a sheet.</summary>
    public void Row(IReadOnlyList<(string Text, double X, double Width, bool RightAlign)> cells,
                    double size = 9, bool bold = false, bool rule = true)
    {
        Room(size * 1.9);
        double baseline = _y - size;
        foreach (var (text, x, width, right) in cells)
        {
            var shown = Fit(text, size, bold, width - 6);
            double at = right ? x + width - 3 - Measure(shown, size, bold) : x + 3;
            Line(shown, at, baseline, size, bold);
        }
        _y -= size * 1.9;
        if (rule) Rule(_margin, _y + size * 0.6, _width - _margin, _y + size * 0.6, 0.4);
    }

    public void Rule(double x1, double y1, double x2, double y2, double thickness = 0.6)
    {
        Write($"{thickness.ToString("0.##", CultureInfo.InvariantCulture)} w 0.75 G\n");
        Write($"{P(x1)} {P(y1)} m {P(x2)} {P(y2)} l S\n");
    }

    /// <summary>Start a new page whether or not this one is full.</summary>
    public void NewPage()
    {
        if (_current.Length > 0) Close();
        _y = _height - _margin;
    }

    private void Room(double needed)
    {
        if (_y - needed >= _margin) return;
        Close();
        _y = _height - _margin;
    }

    private void Close()
    {
        _pages.Add(_current.ToArray());
        _current.SetLength(0);
        _pageCount++;
    }

    private void Line(string text, double x, double y, double size, bool bold)
    {
        if (text.Length == 0) return;
        Write("BT\n");
        Write($"/{(bold ? "FB" : "FR")} {P(size)} Tf\n");
        Write($"{P(x)} {P(y)} Td\n");
        Write($"({Escape(text)}) Tj\n");
        Write("ET\n");
    }

    private void Write(string s) => _current.Write(Encoding.ASCII.GetBytes(s));

    private static string P(double v) => v.ToString("0.###", CultureInfo.InvariantCulture);

    private static string Escape(string text)
    {
        var built = new StringBuilder(text.Length + 8);
        foreach (char c in text)
        {
            if (c is '(' or ')' or '\\') built.Append('\\').Append(c);
            else if (c < 32) built.Append(' ');
            else if (c > 126) built.Append('\\').Append(Convert.ToString(c, 8).PadLeft(3, '0'));
            else built.Append(c);
        }
        return built.ToString();
    }

    // ---------- measuring ----------

    public static double Measure(string text, double size, bool bold)
    {
        double total = 0;
        var widths = bold ? BoldWidths : PlainWidths;
        foreach (char c in text) total += (c >= 32 && c <= 126 ? widths[c - 32] : 556) * size / 1000.0;
        return total;
    }

    private static string Fit(string text, double size, bool bold, double width)
    {
        if (width <= 0) return "";
        if (Measure(text, size, bold) <= width) return text;
        for (int n = text.Length - 1; n > 0; n--)
        {
            var shorter = text[..n] + "…";
            // The ellipsis is not in the base font's set, so measure with three dots and write two.
            if (Measure(text[..n] + "..", size, bold) <= width) return text[..n] + "..";
        }
        return "";
    }

    public static IReadOnlyList<string> Wrap(string text, double size, bool bold, double width)
    {
        var lines = new List<string>();
        if (text.Length == 0) { lines.Add(""); return lines; }

        foreach (var hard in text.Replace("\r\n", "\n").Split('\n'))
        {
            var words = hard.Split(' ');
            var line = new StringBuilder();
            foreach (var word in words)
            {
                var candidate = line.Length == 0 ? word : line + " " + word;
                if (Measure(candidate, size, bold) <= width || line.Length == 0)
                {
                    if (line.Length > 0) line.Append(' ');
                    line.Append(word);
                }
                else { lines.Add(line.ToString()); line.Clear(); line.Append(word); }
            }
            lines.Add(line.ToString());
        }
        return lines;
    }

    // ---------- the file itself ----------

    public byte[] ToBytes()
    {
        if (_current.Length > 0) Close();
        if (_pages.Count == 0) { _pages.Add(Array.Empty<byte>()); _pageCount = 1; }

        var file = new MemoryStream();
        var offsets = new List<long>();
        void Put(string s) => file.Write(Encoding.ASCII.GetBytes(s));
        void Object(int number, string body)
        {
            offsets.Add(file.Position);
            Put($"{number} 0 obj\n{body}\nendobj\n");
        }

        Put("%PDF-1.4\n%âãÏÓ\n");

        int pageObjectsStart = 5;
        var kids = string.Join(" ", Enumerable.Range(0, _pages.Count).Select(i => $"{pageObjectsStart + i * 2} 0 R"));

        Object(1, "<< /Type /Catalog /Pages 2 0 R >>");
        Object(2, $"<< /Type /Pages /Kids [{kids}] /Count {_pages.Count} >>");
        Object(3, "<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica /Encoding /WinAnsiEncoding >>");
        Object(4, "<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica-Bold /Encoding /WinAnsiEncoding >>");

        for (int i = 0; i < _pages.Count; i++)
        {
            int pageNumber = pageObjectsStart + i * 2;
            Object(pageNumber,
                $"<< /Type /Page /Parent 2 0 R /MediaBox [0 0 {P(_width)} {P(_height)}] " +
                $"/Resources << /Font << /FR 3 0 R /FB 4 0 R >> >> /Contents {pageNumber + 1} 0 R >>");

            offsets.Add(file.Position);
            Put($"{pageNumber + 1} 0 obj\n<< /Length {_pages[i].Length} >>\nstream\n");
            file.Write(_pages[i]);
            Put("\nendstream\nendobj\n");
        }

        long startXref = file.Position;
        int count = offsets.Count + 1;
        Put($"xref\n0 {count}\n0000000000 65535 f \n");
        foreach (var offset in offsets) Put($"{offset:D10} 00000 n \n");
        Put($"trailer\n<< /Size {count} /Root 1 0 R >>\nstartxref\n{startXref}\n%%EOF\n");

        return file.ToArray();
    }

    public void Save(string path) => File.WriteAllBytes(path, ToBytes());

    // Widths of the two fonts every reader already has, for characters 32 to 126, in thousandths of the size.
    private static readonly int[] PlainWidths =
    {
        278,278,355,556,556,889,667,191,333,333,389,584,278,333,278,278,
        556,556,556,556,556,556,556,556,556,556,278,278,584,584,584,556,
        1015,667,667,722,722,667,611,778,722,278,500,667,556,833,722,778,
        667,778,722,667,611,722,667,944,667,667,611,278,278,278,469,556,
        333,556,556,500,556,556,278,556,556,222,222,500,222,833,556,556,
        556,556,333,500,278,556,500,722,500,500,500,334,260,334,584,
    };

    private static readonly int[] BoldWidths =
    {
        278,333,474,556,556,889,722,238,333,333,389,584,278,333,278,278,
        556,556,556,556,556,556,556,556,556,556,333,333,584,584,584,611,
        975,722,722,722,722,667,611,778,722,278,556,722,611,833,722,778,
        667,778,722,667,611,722,667,944,667,667,611,333,278,333,584,556,
        333,556,611,556,611,556,333,611,611,278,278,556,278,889,611,611,
        611,611,389,556,333,611,556,778,556,556,500,389,280,389,584,
    };
}
