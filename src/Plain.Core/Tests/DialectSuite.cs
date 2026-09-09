using System.Text;

namespace Plain.Core.Tests;

/// <summary>
/// Office writes the same document two ways. Nearly everything uses the transitional names; Excel's "Strict Open XML"
/// and some other tools use the ISO ones, which spell every element in a different namespace. A reader that knows only
/// one finds nothing and, worse, says nothing: the file opened, and the window was empty.
///
/// Rather than carry a second set of binary fixtures, this rewrites the real ones into the strict names and insists
/// Plain reads them exactly the same.
/// </summary>
public static class DialectSuite
{
    private static readonly (string From, string To)[] Names =
    {
        ("http://schemas.openxmlformats.org/spreadsheetml/2006/main", "http://purl.oclc.org/ooxml/spreadsheetml/main"),
        ("http://schemas.openxmlformats.org/wordprocessingml/2006/main", "http://purl.oclc.org/ooxml/wordprocessingml/main"),
        ("http://schemas.openxmlformats.org/presentationml/2006/main", "http://purl.oclc.org/ooxml/presentationml/main"),
        ("http://schemas.openxmlformats.org/drawingml/2006/main", "http://purl.oclc.org/ooxml/drawingml/main"),
        ("http://schemas.openxmlformats.org/officeDocument/2006/relationships", "http://purl.oclc.org/ooxml/officeDocument/relationships"),
    };

    /// <summary>Rewrite every XML part into the strict names, leaving the package's own plumbing alone.</summary>
    private static byte[] ToStrict(byte[] original)
    {
        var package = OpcPackage.Read(original);
        var parts = new List<(string, byte[])>();
        foreach (var part in package.Parts)
        {
            var content = package.Read(part.Name);
            // The relationship files keep the package namespace even in a strict document; only the markup changes.
            if (part.Name.EndsWith(".xml", StringComparison.OrdinalIgnoreCase) && !part.Name.EndsWith(".rels", StringComparison.OrdinalIgnoreCase))
            {
                var text = OpcPackage.DecodeUtf8(content);
                foreach (var (from, to) in Names) text = text.Replace(from, to);
                content = new UTF8Encoding(false).GetBytes(text);
            }
            parts.Add((part.Name, content));
        }
        return PackageBuilder.Build(parts);
    }

    public static Suite Run()
    {
        var s = new Suite("dialect");
        if (Fixtures.Missing(s, "sheet.xlsx", "doc.docx", "deck.pptx", "book.xlsx")) return s;

        s.Check("the two families are told apart",
            Dialect.Iso.Strict && !Dialect.Transitional.Strict);

        // Excel
        var plainBook = new Workbook(OpcPackage.Open(Fixtures.Path_("sheet.xlsx")!));
        var strictBook = new Workbook(OpcPackage.Read(ToStrict(File.ReadAllBytes(Fixtures.Path_("sheet.xlsx")!))));
        s.Check("a strict workbook is recognised as strict", strictBook.Dialect.Strict);
        s.Equal("a strict workbook finds its sheets", plainBook.Sheets.Count, strictBook.Sheets.Count);
        s.Equal("with the same names", string.Join(',', plainBook.Sheets.Select(x => x.Name)),
                                       string.Join(',', strictBook.Sheets.Select(x => x.Name)));
        s.Equal("text reads the same", plainBook.Sheets[0].Read("A1").Display, strictBook.Sheets[0].Read("A1").Display);
        s.Equal("numbers read the same", plainBook.Sheets[0].Read("B2").Raw, strictBook.Sheets[0].Read("B2").Raw);
        s.Equal("formulas read the same", plainBook.Sheets[0].Read("B5").Formula, strictBook.Sheets[0].Read("B5").Formula);
        s.Equal("number formats apply the same", plainBook.Sheets[0].Read("B2").Display, strictBook.Sheets[0].Read("B2").Display);
        s.Equal("the same cells are found", plainBook.Sheets[0].Cells().Count(), strictBook.Sheets[0].Cells().Count());
        s.Equal("the extent is the same", plainBook.Sheets[0].Extent, strictBook.Sheets[0].Extent);

        // Several sheets, through the strict relationship attribute.
        var strictMany = new Workbook(OpcPackage.Read(ToStrict(File.ReadAllBytes(Fixtures.Path_("book.xlsx")!))));
        s.Equal("a strict workbook finds all three sheets", 3, strictMany.Sheets.Count);
        s.Equal("and reads a later one", "Seat", strictMany.Sheets[1].Read("A2").Display);

        // Word
        var plainDoc = new Document(OpcPackage.Open(Fixtures.Path_("doc.docx")!));
        var strictDoc = new Document(OpcPackage.Read(ToStrict(File.ReadAllBytes(Fixtures.Path_("doc.docx")!))));
        s.Equal("a strict document has the same blocks", plainDoc.BlockCount, strictDoc.BlockCount);
        s.Equal("and the same text", plainDoc.PlainText(), strictDoc.PlainText());
        s.Check("and still knows a heading", strictDoc.Blocks().Any(b => b.IsHeading));
        s.Check("and still knows a table", strictDoc.Blocks().Any(b => b.Kind == BlockKind.TableCell));

        // PowerPoint
        var plainDeck = new Deck(OpcPackage.Open(Fixtures.Path_("deck.pptx")!));
        var strictDeck = new Deck(OpcPackage.Read(ToStrict(File.ReadAllBytes(Fixtures.Path_("deck.pptx")!))));
        s.Equal("a strict deck has the same slides", plainDeck.Slides.Count, strictDeck.Slides.Count);
        s.Equal("and the same titles", plainDeck.Slides[1].Title(), strictDeck.Slides[1].Title());
        s.Equal("and the same lines",
                string.Join('|', plainDeck.Slides[1].Texts().SelectMany(t => t.Lines)),
                string.Join('|', strictDeck.Slides[1].Texts().SelectMany(t => t.Lines)));

        // Editing a strict file must keep it strict, not quietly convert it.
        var strictBytes = ToStrict(File.ReadAllBytes(Fixtures.Path_("sheet.xlsx")!));
        var editable = new Workbook(OpcPackage.Read(strictBytes));
        editable.Sheets[0].Set("A2", "Rewritten");
        editable.Flush();
        var after = new Workbook(OpcPackage.Read(editable.Package.ToBytes()));
        s.Equal("an edit to a strict workbook reads back", "Rewritten", after.Sheets[0].Read("A2").Display);
        s.Check("and the file is still strict", after.Dialect.Strict);
        s.Check("and its untouched cells are intact", after.Sheets[0].Read("A4").Display == "APAC");

        // A workbook whose sheets cannot be found must say so, not show an empty grid.
        var broken = OpcPackage.Read(File.ReadAllBytes(Fixtures.Path_("sheet.xlsx")!));
        broken.WriteText("xl/workbook.xml",
            "<?xml version=\"1.0\"?><workbook xmlns=\"http://example.invalid/not-a-workbook\"><sheets/></workbook>");
        s.Throws<OpcPackage.PackageException>("a workbook with no sheets Plain can find is refused",
            () => new Workbook(OpcPackage.Read(broken.ToBytes())));

        return s;
    }
}
