namespace Plain.Core.Tests;

/// <summary>
/// Page setup: the paper, the way round, the margins, and the line at the top and bottom. The checks that matter
/// most are the ones where a setting is silly, because a margin wider than the paper should give you a page you can
/// still read rather than a file with nothing on it.
/// </summary>
public static class PageSuite
{
    private static string Ascii(byte[] pdf) => System.Text.Encoding.ASCII.GetString(pdf);

    private static PlainFile Handmade(string text)
    {
        var path = System.IO.Path.Combine(System.IO.Path.GetTempPath(), "plain-hand-" + Guid.NewGuid().ToString("N") + ".docx");
        var file = PlainFile.Create(path);
        file.Document!.SetText(0, text);
        return file;
    }

    public static Suite Run()
    {
        var s = new Suite("page");

        // ---- paper and orientation ----
        var a4 = PageSetup.Default;
        s.Check("A4 is taller than it is wide", a4.HeightPoints > a4.WidthPoints);
        s.Check("landscape turns it round", (a4 with { Landscape = true }).WidthPoints > (a4 with { Landscape = true }).HeightPoints);
        s.Check("landscape is the same paper, the other way",
            Math.Abs((a4 with { Landscape = true }).WidthPoints - a4.HeightPoints) < 0.01);

        var letter = a4 with { Paper = "Letter" };
        s.Check("Letter is wider than A4", letter.WidthPoints > a4.WidthPoints);
        s.Check("Letter is shorter than A4", letter.HeightPoints < a4.HeightPoints);
        s.Check("a paper nobody has heard of falls back to A4",
            Math.Abs((a4 with { Paper = "Papyrus" }).WidthPoints - a4.WidthPoints) < 0.01);
        s.Check("the paper name is not case sensitive",
            Math.Abs((a4 with { Paper = "letter" }).WidthPoints - letter.WidthPoints) < 0.01);

        // ---- margins ----
        s.Check("twenty millimetres is about 57 points", Math.Abs(a4.MarginPoints - 56.7) < 1);
        s.Check("a negative margin is not allowed", (a4 with { MarginMm = -10 }).MarginPoints == 0);
        s.Check("an enormous margin is brought back", (a4 with { MarginMm = 500 }).MarginPoints <= 60 * 72.0 / 25.4);
        // 60mm on A4 still leaves a readable column, so that one is left alone; on A5 it does not.
        s.Check("wide margins on big paper are left alone", (a4 with { MarginMm = 60 }).Sensible().MarginMm == 60);
        s.Check("a margin that would leave no room is replaced with one that does",
            (a4 with { Paper = "A5", MarginMm = 60 }).Sensible().MarginMm < 60);
        s.Check("an ordinary margin is left alone", a4.Sensible().MarginMm == a4.MarginMm);

        // ---- what actually comes out ----
        var path = System.IO.Path.Combine(System.IO.Path.GetTempPath(), "plain-page-" + Guid.NewGuid().ToString("N") + ".docx");
        var file = PlainFile.Create(path);
        var doc = file.Document!;
        doc.SetText(0, string.Concat(Enumerable.Repeat(
            "This is a paragraph of ordinary length, put here so the document runs to more than one page. ", 220)));

        var plain = PdfExport.Build(file, "Test", a4 with { Footer = "", Header = "" });
        s.Check("a long document runs to more than one page", plain.Pages > 1);

        var numbered = PdfExport.Build(file, "Test", a4);
        var text = Ascii(numbered.Bytes);
        s.Check("the footer is written", text.Contains("Page 1 of"));
        s.Check("the page number counts up", text.Contains("Page 2 of"));
        s.Check("the total is the real total", text.Contains($"Page 1 of {numbered.Pages}"));
        s.Check("a footer does not cost a page", numbered.Pages == plain.Pages);

        var headed = PdfExport.Build(file, "Test", a4 with { Header = "Woodland Ave" });
        s.Check("a header is written", Ascii(headed.Bytes).Contains("Woodland Ave"));

        var wide = PdfExport.Build(file, "Test", a4 with { Landscape = true });
        s.Check("landscape says so in the paper size",
            Ascii(wide.Bytes).Contains("/MediaBox [0 0 841") || Ascii(wide.Bytes).Contains("/MediaBox [0 0 842"));
        s.Check("landscape fits more on a page, so it takes fewer", wide.Pages <= numbered.Pages);

        var tight = PdfExport.Build(file, "Test", a4 with { MarginMm = 5 });
        s.Check("a small margin fits more on a page", tight.Pages <= numbered.Pages);

        var silly = PdfExport.Build(file, "Test", a4 with { MarginMm = 200 });
        s.Check("a page cannot be all margin", silly.Pages > 0);
        s.Check("and it still holds the words", Ascii(silly.Bytes).Contains("ordinary length"));

        // ---- characters the built-in fonts hold in a different place ----
        // A bullet used to come out as "€42": its Unicode number was written as an octal escape, and a reader takes
        // the first three digits as one character and prints the rest.
        {
            var cpath = System.IO.Path.Combine(System.IO.Path.GetTempPath(), "plain-chars-" + Guid.NewGuid().ToString("N") + ".docx");
            var cfile = PlainFile.Create(cpath);
            cfile.Document!.SetText(0, "A bullet \u2022 a dash \u2014 curly \u201Cquotes\u201D and an ellipsis \u2026");
            var made = PdfExport.Build(cfile, "Characters", PageSetup.Default with { Footer = "" });
            var raw = Ascii(made.Bytes);

            s.Check("a bullet is written where the font keeps it", raw.Contains("\\225"));
            s.Check("an em dash too", raw.Contains("\\227"));
            s.Check("and curly quotes", raw.Contains("\\223") && raw.Contains("\\224"));
            s.Check("and an ellipsis", raw.Contains("\\205"));
            s.Check("nothing is written as an escape a reader would misread", !raw.Contains("\\20042"));
            s.Check("these characters are not reported as unwritable",
                Pdf.CanWrite("\u2022 \u2014 \u201C \u201D \u2026"));
            s.Check("something genuinely outside the fonts still is", !Pdf.CanWrite("\u4E2D"));
            s.Check("and is written as a question mark rather than nonsense",
                Ascii(PdfExport.Build(Handmade("\u4E2D"), "x", PageSetup.Default).Bytes).Contains("?"));
        }

        return s;
    }
}
