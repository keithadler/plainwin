namespace Plain.Core.Tests;

public static class ReplaceSuite
{
    private static string One(string text, string find, string with, bool matchCase = false, bool wholeCell = false)
        => Replace.InText(text, find, with, new Replace.Options(matchCase, wholeCell), out _);

    private static int Count(string text, string find, string with)
    {
        Replace.InText(text, find, with, new Replace.Options(), out int n);
        return n;
    }

    public static Suite Run()
    {
        var s = new Suite("replace");

        s.Equal("the plain case", "hello world", One("hello there", "there", "world"));
        s.Equal("every occurrence", "b b b", One("a a a", "a", "b"));
        s.Equal("counted", 3, Count("a a a", "a", "b"));
        s.Equal("nothing to find changes nothing", "abc", One("abc", "zzz", "!"));
        s.Equal("an empty search is refused", "abc", One("abc", "", "!"));
        s.Equal("replacing with nothing removes it", "ac", One("abc", "b", ""));
        s.Equal("case is ignored by default", "X X", One("a A", "a", "X"));
        s.Equal("until it is not", "X A", One("a A", "a", "X", matchCase: true));
        s.Equal("whole cell only when it is the whole cell", "abc", One("abc", "b", "Z", wholeCell: true));
        s.Equal("whole cell matches the whole cell", "Z", One("abc", "abc", "Z", wholeCell: true));

        // The replacement containing the search term must not send it round for ever.
        s.Equal("a replacement containing the search term terminates", "aa", One("a", "a", "aa"));
        s.Equal("and is counted once", 1, Count("a", "a", "aa"));
        s.Equal("overlapping matches take the first", "xx", One("aaaa", "aa", "x"));

        if (Fixtures.Missing(s, "book.xlsx", "doc.docx", "deck.pptx")) return s;

        // A workbook, across all its sheets.
        var book = Fixtures.Copy("book.xlsx");
        try
        {
            var w = new Workbook(OpcPackage.Open(book));
            var result = Replace.InWorkbook(w, "e", "3", new Replace.Options());
            s.Check("a workbook replace reaches several sheets", result.Cells > 3);
            w.Save(book);

            var again = new Workbook(OpcPackage.Open(book));
            s.Check("the first sheet changed", again.Sheets[0].Read("A2").Display.Contains('3'));
            s.Check("a later sheet changed too", again.Sheets[1].Read("A2").Display.Contains('3'));
        }
        finally { try { File.Delete(book); } catch { } }

        // Replacing text must not disturb numbers or formulas that do not contain it.
        var sheetCopy = Fixtures.Copy("sheet.xlsx");
        try
        {
            var w = new Workbook(OpcPackage.Open(sheetCopy));
            string formulaBefore = w.Sheets[0].Read("B5").Formula!;
            var result = Replace.InWorkbook(w, "EMEA", "Europe", new Replace.Options());
            s.Equal("only the matching cell changed", 1, result.Cells);
            w.Save(sheetCopy);

            var again = new Workbook(OpcPackage.Open(sheetCopy)).Sheets[0];
            s.Equal("the label changed", "Europe", again.Read("A3").Display);
            s.Equal("the formula beside it is untouched", formulaBefore, again.Read("B5").Formula);
            s.Equal("a number is untouched", "412800", again.Read("B2").Raw);
        }
        finally { try { File.Delete(sheetCopy); } catch { } }

        // A formula can be the thing you are replacing in.
        var formulas = Fixtures.Copy("sheet.xlsx");
        try
        {
            var w = new Workbook(OpcPackage.Open(formulas));
            var result = Replace.InWorkbook(w, "SUM", "MAX", new Replace.Options());
            s.Check("formulas are searched too", result.Occurrences >= 1);
            w.Save(formulas);
            s.Check("the formula changed",
                    (new Workbook(OpcPackage.Open(formulas)).Sheets[0].Read("B5").Formula ?? "").Contains("MAX"));

            var w2 = new Workbook(OpcPackage.Open(Fixtures.Path_("sheet.xlsx")!));
            var skipped = Replace.InWorkbook(w2, "SUM", "MAX", new Replace.Options(IncludeFormulas: false));
            s.Equal("and can be left alone when asked", 0, skipped.Occurrences);
        }
        finally { try { File.Delete(formulas); } catch { } }

        // A document.
        var docCopy = Fixtures.Copy("doc.docx");
        try
        {
            var d = new Document(OpcPackage.Open(docCopy));
            var result = Replace.InDocument(d, "Northgate Partners", "Southgate Holdings", new Replace.Options());
            s.Check("a document replace finds the name", result.Occurrences >= 1);
            d.Save(docCopy);
            var text = new Document(OpcPackage.Open(docCopy)).PlainText();
            s.Check("the new name is there", text.Contains("Southgate Holdings"));
            s.Check("the old name is gone", !text.Contains("Northgate Partners"));
            s.Check("the rest of the document survived", text.Contains("Discovery and audit"));
        }
        finally { try { File.Delete(docCopy); } catch { } }

        // A deck.
        var deckCopy = Fixtures.Copy("deck.pptx");
        try
        {
            var deck = new Deck(OpcPackage.Open(deckCopy));
            var result = Replace.InDeck(deck, "nine", "four", new Replace.Options());
            s.Equal("a deck replace finds the word", 1, result.Occurrences);
            deck.Save(deckCopy);
            var after = new Deck(OpcPackage.Open(deckCopy));
            s.Check("the slide line changed",
                    after.Slides[1].Texts().SelectMany(t => t.Lines).Any(l => l.Contains("four working days")));
            s.Equal("the title is untouched", "Where we are today", after.Slides[1].Title());
        }
        finally { try { File.Delete(deckCopy); } catch { } }

        return s;
    }
}
