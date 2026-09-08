using System.Diagnostics;

namespace Plain.Core.Tests;

/// <summary>
/// Guards against the shape of the code going quadratic again. Reading a cell used to scan every row of the sheet, so
/// drawing the bottom of a five thousand row grid walked millions of elements and searching one took twenty seconds.
/// The limits here are deliberately loose: they are meant to catch an algorithm that got worse by a factor of a
/// hundred, not to measure how fast a particular machine is on a particular day.
/// </summary>
public static class ScaleSuite
{
    public static Suite Run()
    {
        var s = new Suite("scale");
        var path = Fixtures.Path_("large.xlsx");
        if (path is null) { s.Check("large fixture found", false, "tests/fixtures/large.xlsx is missing"); return s; }

        var book = new Workbook(OpcPackage.Open(path));
        var sheet = book.Sheets[0];
        var extent = sheet.Extent;
        s.Check($"the sheet really is large ({extent.Row} rows)", extent.Row >= 5000);

        // Warm the sheet so the XML parse is not counted against the reads.
        _ = sheet.Read("A1");

        var sw = Stopwatch.StartNew();
        int last = extent.Row;
        for (int r = last - 25; r < last; r++)
            for (int c = 1; c <= 10; c++)
                _ = sheet.Read(new CellRef(c, r));
        sw.Stop();
        s.Check($"a screen at the bottom draws quickly ({sw.ElapsedMilliseconds} ms)", sw.ElapsedMilliseconds < 400,
                "reading cells near the end of a sheet is scanning the whole sheet");

        sw.Restart();
        int seen = sheet.Cells().Count();
        sw.Stop();
        s.Check($"every cell lists in one pass ({seen} cells, {sw.ElapsedMilliseconds} ms)", sw.ElapsedMilliseconds < 2000,
                "listing cells is quadratic");

        sw.Restart();
        var hits = Search.InWorkbook(book, "Item 4999");
        sw.Stop();
        s.Check($"searching is quick ({sw.ElapsedMilliseconds} ms)", sw.ElapsedMilliseconds < 2000, "search is quadratic");
        s.Check("the search found the row", hits.Count >= 1);

        // The extent is asked for on every repaint, so it must not walk the sheet each time.
        sw.Restart();
        for (int i = 0; i < 500; i++) _ = sheet.Extent;
        sw.Stop();
        s.Check($"the extent is remembered ({sw.ElapsedMilliseconds} ms for 500)", sw.ElapsedMilliseconds < 200,
                "the extent is recomputed on every ask");

        // An edit must invalidate what was remembered.
        sheet.Set(new CellRef(1, extent.Row + 5), "past the end");
        s.Equal("the extent grows after an edit past it", extent.Row + 5, sheet.Extent.Row);

        return s;
    }
}
