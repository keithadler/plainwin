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

        if (Fixtures.Missing(s, "sheet.xlsx", "book.xlsx")) return s;
        var path = Fixtures.Path_("sheet.xlsx")!;

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
            sh.Set("G9", "=1+1");                 // reads nothing at all
            sh.Set("G10", "=SUM(B2:B4)");         // reads B2, which this test changed to 500000
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

            // A total that reads an edited cell must show the new number, not the old one and not the formula.
            s.Equal("a total that reads the edited cell is worked out again", "964700", s2.Read("B5").Raw);
            s.Check("that total keeps its formula", (s2.Read("B5").Formula ?? "").Contains("SUM"));
            s.Equal("a formula's own text survives", "=B2*2", s2.Read("F2").Formula);
            s.Equal("a new formula is worked out", "1000000", s2.Read("F2").Raw);
            s.Equal("a formula that reads nothing is still worked out", "2", s2.Read("G9").Raw);
            s.Equal("and one reading cells the same edit touched", "964700", s2.Read("G10").Raw);
            s.Equal("and shows its number, not itself", "1000000", s2.Read("F2").Display);
            s.Check("a plain number keeps its value", s2.Read("C2").Raw.Length > 0);

            // Cells must stay in ascending order or Excel calls the file damaged.
            var xml = again.Package.ReadText(s2.PartName);
            s.Check("cells stay in order", CellsAreOrdered(xml), "a row lists its cells out of order");
        }
        finally { try { File.Delete(work); } catch { } }

        // A workbook with several sheets must expose all of them, by name, each with its own cells.
        {
            var multi = Fixtures.Path_("book.xlsx")!;
            var many = new Workbook(OpcPackage.Open(multi));
            s.Equal("three sheets", 3, many.Sheets.Count);
            s.Equal("sheets keep their order and names", "Summary,Detail,Notes",
                    string.Join(',', many.Sheets.Select(x => x.Name)));
            s.Equal("the first sheet's cells", "Licences", many.Sheets[0].Read("A2").Display);
            s.Equal("the second sheet's cells", "Seat", many.Sheets[1].Read("A2").Display);
            s.Equal("the third sheet's cells", "Renewal is in March.", many.Sheets[2].Read("A2").Display);
            s.Check("each sheet has its own part",
                    many.Sheets.Select(x => x.PartName).Distinct().Count() == 3);

            // Editing one sheet must leave the others' parts untouched.
            var copy = Fixtures.Copy("book.xlsx");
            try
            {
                var w = new Workbook(OpcPackage.Open(copy));
                var untouchedPart = w.Sheets[2].PartName;
                var before = w.Package.Read(untouchedPart);
                w.Sheets[0].Set("A2", "Subscriptions");
                w.Save(copy);
                var again = new Workbook(OpcPackage.Open(copy));
                s.Equal("the edited sheet changed", "Subscriptions", again.Sheets[0].Read("A2").Display);
                s.Bytes("a sheet nobody edited is byte for byte", before, again.Package.Read(untouchedPart));
                s.Equal("the other sheet still reads", "Seat", again.Sheets[1].Read("A2").Display);
            }
            finally { try { File.Delete(copy); } catch { } }
        }

        Staleness(s);
        Formatting(s);
        NoSharedTable(s);
        // ---- column widths, which sorting and autofit both lean on ----
        if (!Fixtures.Missing(s, "sheet.xlsx"))
        {
            var wf = Fixtures.Copy("sheet.xlsx");
            var wfile = PlainFile.Open(wf);
            var wb = wfile.Workbook!;
            var ws = wb.Sheets[0];

            double was = ws.WidthChars(2);
            ws.SetWidthChars(2, 31.5);
            s.Check("a width set is a width read back", Math.Abs(ws.WidthChars(2) - 31.5) < 0.01);
            s.Check("the column beside it is untouched", Math.Abs(ws.WidthChars(3) - ws.WidthChars(3)) < 0.01);

            // Setting one column inside a range has to split the range, not widen all of it.
            ws.SetWidthChars(5, 12.0);
            ws.SetWidthChars(6, 44.0);
            s.Check("two columns keep their own widths",
                Math.Abs(ws.WidthChars(5) - 12.0) < 0.01 && Math.Abs(ws.WidthChars(6) - 44.0) < 0.01);
            ws.SetWidthChars(5, 9.0);
            s.Check("changing one again leaves the other alone", Math.Abs(ws.WidthChars(6) - 44.0) < 0.01);

            ws.SetWidthChars(7, 100000);
            s.Check("a silly width is brought back into range", ws.WidthChars(7) <= 255);
            ws.SetWidthChars(0, 20);
            ws.SetWidthChars(20000, 20);
            s.Check("a column outside the sheet is refused quietly", ws.WidthChars(1) > 0);

            wfile.Flush();   // the models hold the change until asked; the package is only bytes
            var written = PlainFile.Read(wfile.Package.ToBytes(), wf);
            var reread = written.Workbook!.Sheets[0];
            s.Check("the widths survive a save", Math.Abs(reread.WidthChars(2) - 31.5) < 0.01);
            s.Check("and so does the one beside it", Math.Abs(reread.WidthChars(6) - 44.0) < 0.01);
            s.Check("the file still round trips after a width change",
                PlainFile.Read(written.Package.ToBytes(), wf).Workbook is not null);
            s.Check("only the sheet was rewritten, the rest kept byte for byte",
                wfile.Package.Parts.Count(pp => pp.Edited) <= 2);

            s.Check("autofit asks for room for the longest value", ws.WidestChars(1) > 0);
        }

        // ---- frozen header rows ----
        if (!Fixtures.Missing(s, "sheet.xlsx"))
        {
            var ff = Fixtures.Copy("sheet.xlsx");
            var ffile = PlainFile.Open(ff);
            var fs = ffile.Workbook!.Sheets[0];

            fs.SetFrozenRows(1);
            s.Check("one row frozen reads back as one", fs.FrozenRows == 1);
            fs.SetFrozenRows(4);
            s.Check("changing it does not stack up panes", fs.FrozenRows == 4);
            fs.SetFrozenRows(0);
            s.Check("none means none", fs.FrozenRows == 0);

            fs.SetFrozenRows(2);
            ffile.Flush();
            var back = PlainFile.Read(ffile.Package.ToBytes(), ff).Workbook!.Sheets[0];
            s.Check("freezing survives a save", back.FrozenRows == 2);
            s.Check("the file still round trips with a frozen pane",
                PlainFile.Read(ffile.Package.ToBytes(), ff).Workbook is not null);
        }

        // ---- joined blocks ----
        {
            var corpus = Environment.GetEnvironmentVariable("PLAIN_CORPUS");
            var demo = corpus is not null ? System.IO.Path.Combine(corpus, "quarter.xlsx") : null;
            if (demo is not null && File.Exists(demo))
            {
                var ms = PlainFile.Open(demo).Workbook!.Sheets[0];
                s.Check("a joined heading is found", ms.Merges.Count >= 1);
                s.Check("it covers more than one cell",
                    ms.Merges[0].To.Column > ms.Merges[0].From.Column || ms.Merges[0].To.Row > ms.Merges[0].From.Row);
                s.Check("a cell inside it belongs to it", ms.MergeAt(new CellRef(3, 1)) is not null);
                s.Check("the corner belongs to it too", ms.MergeAt(new CellRef(1, 1)) is not null);
                s.Check("and it names the corner that holds the value",
                    ms.MergeAt(new CellRef(3, 1))!.Value.From == new CellRef(1, 1));
                s.Check("a cell outside it does not", ms.MergeAt(new CellRef(1, 5)) is null);
            }
            else s.Check("no demo folder set (PLAIN_CORPUS); skipping the joined block checks", true);

            if (!Fixtures.Missing(s, "sheet.xlsx"))
            {
                var plainSheet = PlainFile.Open(Fixtures.Copy("sheet.xlsx")).Workbook!.Sheets[0];
                s.Check("a sheet with nothing joined says so", plainSheet.Merges.Count == 0);
                s.Check("and no cell belongs to a block", plainSheet.MergeAt(new CellRef(1, 1)) is null);
            }
        }

        return s;
    }

    /// <summary>
    /// A workbook with no shared table of text must still take text. Three of thirty-nine real workbooks had none,
    /// and Plain refused to type into any of them; the words go straight into the cell instead, which adds no part.
    /// </summary>
    private static void NoSharedTable(Suite s)
    {
        var work = Fixtures.Copy("sheet.xlsx");
        try
        {
            // Take the shared table out of the package to make a workbook of the awkward kind.
            var pkg = OpcPackage.Open(work);
            var contentTypes = pkg.ReadText("[Content_Types].xml")
                .Replace("<Override PartName=\"/xl/sharedStrings.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml\"/>", "");
            pkg.WriteText("[Content_Types].xml", contentTypes);
            pkg.WriteText("xl/sharedStrings.xml", "");     // present but empty: Plain must treat it as absent
            var stripped = OpcPackage.Read(pkg.ToBytes());

            var book = new Workbook(stripped);
            var sheet = book.Sheets[0];
            s.Check("a workbook with no usable shared table says so", !book.HasSharedStrings);

            sheet.Set("A9", "typed without a shared table");
            sheet.Set("B9", "1234");
            book.Flush();
            var saved = OpcPackage.Read(stripped.ToBytes());
            var again = new Workbook(saved).Sheets[0];
            s.Equal("text lands in the cell itself", "typed without a shared table", again.Read("A9").Display);
            s.Equal("and a number is still a number", "1234", again.Read("B9").Raw);
            s.Check("the text is written inline",
                    saved.ReadText(again.PartName).Contains("inlineStr"),
                    "text was not written into the cell");
        }
        finally { try { File.Delete(work); } catch { } }
    }

    /// <summary>Changing how a cell shows its number must not change the number, or anything else about the cell.</summary>
    private static void Formatting(Suite s)
    {
        var work = Fixtures.Copy("sheet.xlsx");
        try
        {
            var w = new Workbook(OpcPackage.Open(work));
            var sheet = w.Sheets[0];
            string raw = sheet.Read("B2").Raw;

            sheet.SetFormat(new[] { CellRef.Parse("B2") }, "0.00%");
            s.Equal("the stored number is untouched", raw, sheet.Read("B2").Raw);
            s.Check("the cell now shows as a percentage", sheet.Read("B2").Display.EndsWith("%"));
            s.Equal("and reports its format", "0.00%", sheet.FormatOf(CellRef.Parse("B2")));

            w.Save(work);
            var again = new Workbook(OpcPackage.Open(work)).Sheets[0];
            s.Equal("the format survives a save", "0.00%", again.FormatOf(CellRef.Parse("B2")));
            s.Equal("and so does the number", raw, again.Read("B2").Raw);
            s.Check("a neighbouring cell is unaffected", !again.Read("C2").Display.EndsWith("%"));

            // Asking for the same format twice must not grow the style table for ever.
            var w2 = new Workbook(OpcPackage.Open(work));
            var s2 = w2.Sheets[0];
            s2.SetFormat(new[] { CellRef.Parse("C2") }, "0.00%");
            s2.SetFormat(new[] { CellRef.Parse("D2") }, "0.00%");
            s.Equal("cells asking for one format share one style",
                    s2.FormatOf(CellRef.Parse("C2")), s2.FormatOf(CellRef.Parse("D2")));

            // Putting it back to General.
            s2.SetFormat(new[] { CellRef.Parse("C2") }, "");
            s.Equal("clearing the format clears it", "", s2.FormatOf(CellRef.Parse("C2")));

            // A whole block at once, which is how anyone actually formats a column.
            var block = Enumerable.Range(2, 3).Select(r => new CellRef(2, r)).ToList();
            s2.SetFormat(block, "#,##0");
            s.Check("every cell in the block took the format",
                    block.All(c => s2.FormatOf(c) == "#,##0"));
        }
        finally { try { File.Delete(work); } catch { } }
    }

    /// <summary>Only the totals that read an edited cell may lose their value; the rest must keep theirs.</summary>
    private static void Staleness(Suite s)
    {
        var work = Fixtures.Copy("sheet.xlsx");
        try
        {
            // B5 = SUM(B2:B4); C5 and D5 total their own columns; E2..E4 divide within their row.
            var w = new Workbook(OpcPackage.Open(work));
            var sheet = w.Sheets[0];
            string cachedC5 = sheet.Read("C5").Raw, cachedD5 = sheet.Read("D5").Raw, cachedE3 = sheet.Read("E3").Raw;
            s.Check("the fixture starts with cached totals", cachedC5.Length > 0 && cachedD5.Length > 0);

            sheet.Set("B2", "500000");
            w.Save(work);

            var again = new Workbook(OpcPackage.Open(work)).Sheets[0];
            s.Equal("the total over the edited column is worked out again", "964700", again.Read("B5").Raw);
            s.Equal("a total over another column keeps its value", cachedC5, again.Read("C5").Raw);
            s.Equal("a second untouched total keeps its value", cachedD5, again.Read("D5").Raw);
            s.Equal("a formula in an untouched row keeps its value", cachedE3, again.Read("E3").Raw);
            s.Check("the untouched totals still show a number", again.Read("C5").Display.Length > 0);
        }
        finally { try { File.Delete(work); } catch { } }

        // Staleness must travel: a total of a total is stale too.
        var chain = Fixtures.Copy("sheet.xlsx");
        try
        {
            var w = new Workbook(OpcPackage.Open(chain));
            var sheet = w.Sheets[0];
            sheet.Set("H1", "10");
            sheet.Set("H2", "=H1*2");
            sheet.Set("H3", "=H2+1");
            w.Save(chain);

            var w2 = new Workbook(OpcPackage.Open(chain));
            var s2 = w2.Sheets[0];
            s2.Set("H1", "20");
            w2.Save(chain);

            var after = new Workbook(OpcPackage.Open(chain)).Sheets[0];
            s.Equal("a formula reading the edit is worked out", "40", after.Read("H2").Raw);
            s.Equal("a formula reading that formula follows it", "41", after.Read("H3").Raw);
            s.Equal("the edited value itself is kept", "20", after.Read("H1").Raw);
        }
        finally { try { File.Delete(chain); } catch { } }

        // An edit on one sheet must clear a total on another that reads it, and nothing else.
        var multi = Fixtures.Copy("book.xlsx");
        try
        {
            var w = new Workbook(OpcPackage.Open(multi));
            w.Sheets[0].Set("D1", "=SUM(Detail!B2:B4)");
            w.Sheets[0].Set("D2", "=SUM(Notes!A1:A3)");
            w.Save(multi);

            var w2 = new Workbook(OpcPackage.Open(multi));
            w2.Sheets[1].Set("B2", "999");
            w2.Save(multi);

            var after = new Workbook(OpcPackage.Open(multi)).Sheets[0];
            s.Check("a cross-sheet total over the edit is worked out again",
                    after.Read("D1").Raw.Length > 0 && after.Read("D1").Raw != "0");
            s.Check("a cross-sheet total over another sheet is untouched", after.Read("D2").Formula is not null);
        }
        finally { try { File.Delete(multi); } catch { } }
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
