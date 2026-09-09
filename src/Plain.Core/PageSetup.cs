namespace Plain.Core;

/// <summary>
/// How a page is laid out when Plain prints or writes a PDF: the paper, which way round it is, how much white
/// there is at the edges, and the line at the top and bottom.
///
/// This is Plain's own page, not Word's. Plain does not know where Word would break a page and does not pretend to,
/// so what comes out is the text as Plain shows it, laid out on paper you chose. Saying that plainly is better than
/// producing something that looks like a facsimile and is not one.
/// </summary>
public sealed record PageSetup
{
    /// <summary>Papers by the names people ask for them by, in points, portrait.</summary>
    public static readonly (string Name, double Width, double Height)[] Papers =
    {
        ("A4", 595.28, 841.89),
        ("Letter", 612, 792),
        ("Legal", 612, 1008),
        ("A3", 841.89, 1190.55),
        ("A5", 419.53, 595.28),
    };

    public string Paper { get; init; } = "A4";
    public bool Landscape { get; init; }

    /// <summary>The white at every edge, in millimetres, because that is how people say it.</summary>
    public double MarginMm { get; init; } = 20;

    /// <summary>Lines at the top and bottom. {page} and {pages} are filled in. Empty means none.</summary>
    public string Header { get; init; } = "";
    public string Footer { get; init; } = "Page {page} of {pages}";

    public static PageSetup Default => new();

    private (double W, double H) Sheet()
    {
        var found = Papers.FirstOrDefault(p => p.Name.Equals(Paper, StringComparison.OrdinalIgnoreCase));
        if (found.Name is null) found = Papers[0];
        return Landscape ? (found.Height, found.Width) : (found.Width, found.Height);
    }

    public double WidthPoints => Sheet().W;
    public double HeightPoints => Sheet().H;

    /// <summary>A millimetre is 72/25.4 points. Kept sane so a margin can never swallow the whole page.</summary>
    public double MarginPoints => Math.Clamp(MarginMm, 0, 60) * 72.0 / 25.4;

    /// <summary>A page with nothing left to write on is not a page; keep at least a usable column of text.</summary>
    public PageSetup Sensible()
    {
        var (w, h) = Sheet();
        double margin = MarginPoints;
        if (w - 2 * margin >= 120 && h - 2 * margin >= 160) return this;
        return this with { MarginMm = 15 };
    }

    /// <summary>Build a page ready to be written on.</summary>
    public Pdf NewPdf()
    {
        var ok = Sensible();
        return new Pdf(ok.WidthPoints, ok.HeightPoints, ok.MarginPoints)
        {
            Header = ok.Header,
            Footer = ok.Footer,
        };
    }
}
