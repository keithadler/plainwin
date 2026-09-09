namespace Plain.Core.Tests;

public static class GridSuite
{
    private static string Adjust(string formula, GridEdit edit, int at) => Grid.Adjust(formula, edit, at, null);

    public static Suite Run()
    {
        var s = new Suite("grid");

        // Where a cell lands.
        s.Equal("a row below the insertion moves down", new CellRef(1, 6), Grid.Move(new CellRef(1, 5), GridEdit.InsertRow, 3));
        s.Equal("a row above it stays", new CellRef(1, 2), Grid.Move(new CellRef(1, 2), GridEdit.InsertRow, 3));
        s.Equal("the row at the insertion moves down", new CellRef(1, 4), Grid.Move(new CellRef(1, 3), GridEdit.InsertRow, 3));
        s.Check("the deleted row is gone", Grid.Move(new CellRef(1, 3), GridEdit.DeleteRow, 3) is null);
        s.Equal("a row below a deletion moves up", new CellRef(1, 4), Grid.Move(new CellRef(1, 5), GridEdit.DeleteRow, 3));
        s.Equal("a column right of an insertion moves right", new CellRef(4, 1), Grid.Move(new CellRef(3, 1), GridEdit.InsertColumn, 3));
        s.Check("the deleted column is gone", Grid.Move(new CellRef(3, 1), GridEdit.DeleteColumn, 3) is null);

        // Formulas must still mean what they meant.
        s.Equal("a total grows when a row goes inside it", "=SUM(B2:B5)", Adjust("=SUM(B2:B4)", GridEdit.InsertRow, 3));
        s.Equal("a total grows when a row goes in at its end", "=SUM(B2:B5)", Adjust("=SUM(B2:B4)", GridEdit.InsertRow, 4));
        s.Equal("a total below the insertion moves whole", "=SUM(B6:B8)", Adjust("=SUM(B5:B7)", GridEdit.InsertRow, 3));
        s.Equal("a total above the insertion is untouched", "=SUM(B2:B4)", Adjust("=SUM(B2:B4)", GridEdit.InsertRow, 9));
        s.Equal("a total shrinks when a row inside it goes", "=SUM(B2:B3)", Adjust("=SUM(B2:B4)", GridEdit.DeleteRow, 3));
        s.Equal("a single cell moves", "=B6", Adjust("=B5", GridEdit.InsertRow, 3));
        s.Equal("a deleted single cell is an error, not a wrong answer", "=#REF!", Adjust("=B3", GridEdit.DeleteRow, 3));
        s.Equal("a range that loses everything is an error", "=SUM(#REF!)", Adjust("=SUM(B3:B3)", GridEdit.DeleteRow, 3));

        // Columns behave the same way round.
        s.Equal("a total grows across a new column", "=SUM(B2:E2)", Adjust("=SUM(B2:D2)", GridEdit.InsertColumn, 3));
        s.Equal("a total shrinks when a column goes", "=SUM(B2:C2)", Adjust("=SUM(B2:D2)", GridEdit.DeleteColumn, 3));

        // Dollars must survive, or a pinned reference quietly stops being pinned.
        s.Equal("a fixed row keeps its dollar", "=$B$6", Adjust("=$B$5", GridEdit.InsertRow, 3));
        s.Equal("a half fixed reference keeps its half", "=B$6", Adjust("=B$5", GridEdit.InsertRow, 3));
        s.Equal("a fixed range keeps both", "=SUM($B$2:$B$5)", Adjust("=SUM($B$2:$B$4)", GridEdit.InsertRow, 3));

        // Everything that is not a reference must come through untouched.
        s.Equal("text in the formula survives", "=IF(B6>0,\"yes\",\"no\")", Adjust("=IF(B5>0,\"yes\",\"no\")", GridEdit.InsertRow, 3));
        s.Equal("function names are not references", "=ROUND(B6,2)", Adjust("=ROUND(B5,2)", GridEdit.InsertRow, 3));
        s.Equal("a formula with no references is unchanged", "=1+2", Adjust("=1+2", GridEdit.InsertRow, 3));
        s.Equal("spacing is preserved", "=SUM( B2:B5 )", Adjust("=SUM( B2:B4 )", GridEdit.InsertRow, 3));

        // A reference to another sheet only moves when that sheet is the one changing.
        s.Equal("another sheet is left alone", "=SUM(Other!B2:B4)",
                Grid.Adjust("=SUM(Other!B2:B4)", GridEdit.InsertRow, 3, "This"));
        s.Equal("the named sheet does move", "=SUM(Other!B2:B5)",
                Grid.Adjust("=SUM(Other!B2:B4)", GridEdit.InsertRow, 3, "Other"));
        s.Equal("a quoted sheet name survives the rewrite", "=SUM('My Sheet'!B2:B5)",
                Grid.Adjust("=SUM('My Sheet'!B2:B4)", GridEdit.InsertRow, 3, "My Sheet"));

        OnRealSheets(s);
        return s;
    }

    /// <summary>The whole thing, on a real workbook: cells move, formulas follow, and the file still opens.</summary>
    private static void OnRealSheets(Suite s)
    {
        if (Fixtures.Missing(s, "sheet.xlsx", "book.xlsx")) return;

        var work = Fixtures.Copy("sheet.xlsx");
        try
        {
            var w = new Workbook(OpcPackage.Open(work));
            var sheet = w.Sheets[0];
            // Rows: 1 headers, 2 North America, 3 EMEA, 4 APAC, 5 Total =SUM(B2:B4)
            s.Equal("before: the total sums three rows", "=SUM(B2:B4)", sheet.Read("B5").Formula);

            w.Apply(sheet, GridEdit.InsertRow, 3);
            s.Equal("the row that was third moved down", "EMEA", sheet.Read("A4").Display);
            s.Equal("the row above it stayed", "North America", sheet.Read("A2").Display);
            s.Check("the new row is empty", sheet.Read("A3").IsEmpty);
            s.Equal("the total moved down with its row", "APAC", sheet.Read("A5").Display);
            s.Equal("and now sums four rows", "=SUM(B2:B5)", sheet.Read("B6").Formula);

            w.Save(work);
            var again = new Workbook(OpcPackage.Open(work)).Sheets[0];
            s.Equal("it survives a save", "=SUM(B2:B5)", again.Read("B6").Formula);
            s.Equal("and the moved label is there", "EMEA", again.Read("A4").Display);
            s.Check("cells are still in order", CellsAreOrdered(again));
        }
        finally { try { File.Delete(work); } catch { } }

        var deleting = Fixtures.Copy("sheet.xlsx");
        try
        {
            var w = new Workbook(OpcPackage.Open(deleting));
            var sheet = w.Sheets[0];
            w.Apply(sheet, GridEdit.DeleteRow, 3);      // remove EMEA
            s.Equal("the row below moved up", "APAC", sheet.Read("A3").Display);
            s.Equal("the total came up too", "Total", sheet.Read("A4").Display);
            s.Equal("and now sums two rows", "=SUM(B2:B3)", sheet.Read("B4").Formula);
            s.Check("the deleted label is gone", !sheet.Cells().Any(c => c.Display == "EMEA"));

            w.Save(deleting);
            s.Equal("it survives a save", "=SUM(B2:B3)",
                    new Workbook(OpcPackage.Open(deleting)).Sheets[0].Read("B4").Formula);
        }
        finally { try { File.Delete(deleting); } catch { } }

        var columns = Fixtures.Copy("sheet.xlsx");
        try
        {
            var w = new Workbook(OpcPackage.Open(columns));
            var sheet = w.Sheets[0];
            string q1 = sheet.Read("B1").Display;
            w.Apply(sheet, GridEdit.InsertColumn, 2);
            s.Equal("the column moved right", q1, sheet.Read("C1").Display);
            s.Check("the new column is empty", sheet.Read("B1").IsEmpty);
            s.Equal("a formula followed its column", "=SUM(C2:C4)", sheet.Read("C5").Formula);
            s.Equal("the labels stayed put", "Region", sheet.Read("A1").Display);
        }
        finally { try { File.Delete(columns); } catch { } }

        // A formula on another sheet that reads the changed one must follow it.
        var across = Fixtures.Copy("book.xlsx");
        try
        {
            var w = new Workbook(OpcPackage.Open(across));
            w.Sheets[0].Set("D1", "=SUM(Detail!B2:B4)");
            w.Apply(w.Sheets[1], GridEdit.InsertRow, 3);
            s.Equal("a formula on another sheet follows the change", "=SUM(Detail!B2:B5)", w.Sheets[0].Read("D1").Formula);

            w.Save(across);
            s.Equal("and survives a save", "=SUM(Detail!B2:B5)",
                    new Workbook(OpcPackage.Open(across)).Sheets[0].Read("D1").Formula);
        }
        finally { try { File.Delete(across); } catch { } }

        // A row past the end, and a nonsense row number.
        var edges = Fixtures.Copy("sheet.xlsx");
        try
        {
            var w = new Workbook(OpcPackage.Open(edges));
            s.Throws<ArgumentOutOfRangeException>("row zero is refused", () => w.Apply(w.Sheets[0], GridEdit.InsertRow, 0));
            s.Throws<ArgumentOutOfRangeException>("past the last row is refused", () => w.Apply(w.Sheets[0], GridEdit.InsertRow, 2_000_000));
            w.Apply(w.Sheets[0], GridEdit.InsertRow, 900);
            s.Check("a row inserted past the data changes nothing visible", w.Sheets[0].Read("A2").Display == "North America");
        }
        finally { try { File.Delete(edges); } catch { } }
    }

    private static bool CellsAreOrdered(Sheet sheet)
    {
        int lastRow = 0;
        foreach (var group in sheet.Cells().GroupBy(c => c.Ref.Row))
        {
            if (group.Key < lastRow) return false;
            lastRow = group.Key;
            int lastColumn = 0;
            foreach (var cell in group)
            {
                if (cell.Ref.Column <= lastColumn) return false;
                lastColumn = cell.Ref.Column;
            }
        }
        return true;
    }
}
