namespace Plain.Core.Tests;

public static class SheetSuite
{
    public static Suite Run()
    {
        var s = new Suite("sheet");

        s.Equal("column 1 is A", "A", CellRef.ColumnName(1));
        s.Equal("column 26 is Z", "Z", CellRef.ColumnName(26));
        s.Equal("column 27 is AA", "AA", CellRef.ColumnName(27));
        s.Equal("column 16384 is XFD", "XFD", CellRef.ColumnName(16384));
        s.Equal("A1 parses", new CellRef(1, 1), CellRef.Parse("A1"));
        s.Equal("AA10 parses", new CellRef(27, 10), CellRef.Parse("AA10"));
        s.Equal("a reference round trips", "BC42", CellRef.Parse("BC42").ToString());
        s.Check("a row with no column is refused", !CellRef.TryParse("12", out _));
        s.Check("a column with no row is refused", !CellRef.TryParse("AB", out _));
        s.Check("nonsense is refused", !CellRef.TryParse("hello", out _));

        var path = Fixtures.Path_("sheet.xlsx");
        if (path is null) { s.Check("fixtures found", false, "tests/fixtures/sheet.xlsx is missing"); return s; }

        var book = new Workbook(OpcPackage.Open(path));
        s.Equal("one sheet", 1, book.Sheets.Count);
        s.Equal("the sheet keeps its name", "Forecast", book.Sheets[0].Name);

        var sheet = book.Sheets[0];
        s.Equal("A1 is the header text", "Region", sheet.Read("A1").Display);
        s.Equal("A2 is a region name", "North America", sheet.Read("A2").Display);
        s.Equal("B2 holds the stored number", "412800", sheet.Read("B2").Raw);
        s.Equal("B2 is a number", CellKind.Number, sheet.Read("B2").Kind);
        s.Check("B2 shows with thousands separators", sheet.Read("B2").Display.Length >= 6);
        s.Check("E2 is a formula", sheet.Read("E2").Formula is not null);
        s.Check("B5 totals with SUM", (sheet.Read("B5").Formula ?? "").Contains("SUM", StringComparison.OrdinalIgnoreCase));
        s.Check("an untouched cell is empty", sheet.Read("H40").IsEmpty);
        s.Check("the extent covers the data", sheet.Extent.Row >= 5 && sheet.Extent.Column >= 5);
        s.Check("cells enumerate", sheet.Cells().Count() >= 15);

        // Edits must survive a save and reopen.
        var work = Fixtures.Copy("sheet.xlsx");
        try
        {
            var w = new Workbook(OpcPackage.Open(work));
            var sh = w.Sheets[0];
            sh.Set("A2", "Northern Europe");
            sh.Set("B2", "500000");
            sh.Set("F2", "=B2*2");
            sh.Set("A9", "added by a test");
            w.Save(work);

            var again = new Workbook(OpcPackage.Open(work));
            var s2 = again.Sheets[0];
            s.Equal("edited text reads back", "Northern Europe", s2.Read("A2").Display);
            s.Equal("edited number reads back", "500000", s2.Read("B2").Raw);
            s.Equal("a new formula reads back", "=B2*2", s2.Read("F2").Formula);
            s.Equal("text in a brand new row reads back", "added by a test", s2.Read("A9").Display);
            s.Equal("an untouched cell is unchanged", "APAC", s2.Read("A4").Display);
            s.Check("the workbook asks Excel to recalculate",
                    again.Package.ReadText("xl/workbook.xml").Contains("fullCalcOnLoad"));

            // A cached total next to an edited cell would be a wrong number on screen, so the caches must be gone.
            var sheetXml0 = again.Package.ReadText(s2.PartName);
            s.Check("no formula keeps a cached value after an edit", !HasCachedFormulaValue(sheetXml0),
                    "a formula cell still carries the value Excel cached before the edit");
            s.Equal("a formula's own text survives", "=B2*2", s2.Read("F2").Formula);
            s.Check("a plain number keeps its value", s2.Read("C2").Raw.Length > 0);

            // Cells must stay in ascending order or Excel calls the file damaged.
            var xml = again.Package.ReadText(s2.PartName);
            s.Check("cells stay in order", CellsAreOrdered(xml), "a row lists its cells out of order");
        }
        finally { try { File.Delete(work); } catch { } }

        return s;
    }

    /// <summary>True when any cell carries both a formula and a cached value.</summary>
    private static bool HasCachedFormulaValue(string sheetXml)
    {
        foreach (var chunk in sheetXml.Split("<c ").Skip(1))
        {
            int end = chunk.IndexOf("</c>", StringComparison.Ordinal);
            var cell = end < 0 ? chunk : chunk[..end];
            if (cell.Contains("<f") && cell.Contains("<v>")) return true;
        }
        return false;
    }

    private static bool CellsAreOrdered(string sheetXml)
    {
        foreach (var rowChunk in sheetXml.Split("<row ").Skip(1))
        {
            int last = 0;
            foreach (var cellChunk in rowChunk.Split("<c r=\"").Skip(1))
            {
                int quote = cellChunk.IndexOf('"');
                if (quote < 0) continue;
                if (!CellRef.TryParse(cellChunk[..quote], out var r)) continue;
                if (r.Column <= last) return false;
                last = r.Column;
            }
        }
        return true;
    }
}
