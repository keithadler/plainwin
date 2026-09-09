namespace Plain.Core.Tests;

/// <summary>
/// Adding, renaming, moving and removing sheets, and the formatting that goes with a cell: where its contents sit,
/// what colour it wears, how tall its row is and what stays on screen.
///
/// Renaming is the one with teeth. A formula saying Detail!B2 means the sheet called Detail, so a rename that does
/// not rewrite the formulas naming it leaves them pointing at nothing. Several checks here exist only for that.
/// </summary>
public static class SheetsSuite
{
    private static PlainFile Fresh()
    {
        var path = System.IO.Path.Combine(System.IO.Path.GetTempPath(), "plain-sheets-" + Guid.NewGuid().ToString("N") + ".xlsx");
        return PlainFile.Create(path);
    }

    private static PlainFile Reopen(PlainFile file)
    {
        file.Flush();
        return PlainFile.Read(file.Package.ToBytes(), file.Path);
    }

    public static Suite Run()
    {
        var s = new Suite("sheets");

        // ---- names it will not take ----
        {
            var book = Fresh().Workbook!;
            s.Check("an empty name is refused", Sheets.CheckName(book, "  ") is Sheets.Refused);
            s.Check("a very long name is refused", Sheets.CheckName(book, new string('x', 32)) is Sheets.Refused);
            s.Check("a name with a slash is refused", Sheets.CheckName(book, "in/out") is Sheets.Refused);
            s.Check("a name with brackets is refused", Sheets.CheckName(book, "a[b]") is Sheets.Refused);
            s.Check("a name starting with an apostrophe is refused", Sheets.CheckName(book, "'odd") is Sheets.Refused);
            s.Check("a name already taken is refused", Sheets.CheckName(book, book.Sheets[0].Name) is Sheets.Refused);
            s.Check("the same name in different case is still taken",
                Sheets.CheckName(book, book.Sheets[0].Name.ToUpperInvariant()) is Sheets.Refused);
            s.Check("an ordinary name is fine", Sheets.CheckName(book, "Detail") is Sheets.Done);
        }

        // ---- adding ----
        {
            var file = Fresh();
            var book = file.Workbook!;
            int was = book.Sheets.Count;
            s.Check("a sheet can be added", Sheets.Add(book, "Detail", was) is Sheets.Done);
            s.Check("there is one more", book.Sheets.Count == was + 1);
            s.Check("and it has the name asked for", book.Sheets[^1].Name == "Detail");

            var back = Reopen(file);
            s.Check("the workbook still opens", back.Workbook is not null);
            s.Check("and still has the new sheet", back.Workbook!.Sheets.Any(x => x.Name == "Detail"));
            s.Check("which can be written to",
                Threw(() => back.Workbook.Sheets.First(x => x.Name == "Detail").Set("A1", "hello")) == false);

            var again = Reopen(back);
            s.Check("and what was written is there", again.Workbook!.Sheets.First(x => x.Name == "Detail").Read("A1").Display == "hello");
        }

        // ---- renaming, and the formulas that named it ----
        {
            var file = Fresh();
            var book = file.Workbook!;
            Sheets.Add(book, "Detail", 1);
            var detail = book.Sheets.First(x => x.Name == "Detail");
            detail.Set("B2", "40");
            book.Sheets[0].Set("A1", "=Detail!B2*2");
            book.Sheets[0].Set("A2", "=SUM(Detail!B1:B9)");
            book.Sheets[0].Set("A3", "=1+1");

            var done = Sheets.Rename(book, book.Sheets.ToList().IndexOf(detail), "Invoices");
            s.Check("renaming says it did", done is Sheets.Done);
            s.Check("the sheet has the new name", detail.Name == "Invoices");
            s.Check("a formula that named it now names the new one",
                book.Sheets[0].Read("A1").Formula!.Contains("Invoices"));
            s.Check("and no longer names the old one",
                !book.Sheets[0].Read("A1").Formula!.Contains("Detail"));
            s.Check("a range that named it follows too",
                book.Sheets[0].Read("A2").Formula!.Contains("Invoices"));
            s.Check("a formula that named nothing is untouched",
                book.Sheets[0].Read("A3").Formula == "=1+1");
            // A rename writes the formula back into the file, where it is stored without its leading sign.
            // Putting one back that already had it would give the cell "==Invoices!B2".
            s.Check("a rewritten formula has exactly one equals sign",
                book.Sheets[0].Read("A1").Formula!.StartsWith("=") && !book.Sheets[0].Read("A1").Formula!.StartsWith("=="));
            s.Check("and it says how many it changed", done is Sheets.Done d && d.What.Contains("2"));

            var back = Reopen(file);
            s.Check("the rename survives a save", back.Workbook!.Sheets.Any(x => x.Name == "Invoices"));
            s.Check("and so do the rewritten formulas",
                back.Workbook.Sheets[0].Read("A1").Formula!.Contains("Invoices"));

            s.Check("renaming to a name already taken is refused",
                Sheets.Rename(back.Workbook, 0, "Invoices") is Sheets.Refused);
        }

        // ---- moving ----
        {
            var file = Fresh();
            var book = file.Workbook!;
            Sheets.Add(book, "Second", 1);
            Sheets.Add(book, "Third", 2);
            var order = string.Join(",", book.Sheets.Select(x => x.Name));

            s.Check("a sheet can be moved", Sheets.Move(book, 1, 3) is Sheets.Done);
            s.Check("and the order really changed", string.Join(",", book.Sheets.Select(x => x.Name)) != order);

            var back = Reopen(file);
            s.Check("the new order survives a save",
                string.Join(",", back.Workbook!.Sheets.Select(x => x.Name)) == string.Join(",", book.Sheets.Select(x => x.Name)));
            s.Check("moving one that is not there is refused", Sheets.Move(book, 99, 1) is Sheets.Refused);
        }

        // ---- removing, and the refusal that matters ----
        {
            var file = Fresh();
            var book = file.Workbook!;
            Sheets.Add(book, "Detail", 1);
            book.Sheets[0].Set("A1", "=Detail!B2");
            s.Check("a sheet a formula reads cannot be taken out",
                Sheets.Remove(book, book.Sheets.ToList().FindIndex(x => x.Name == "Detail")) is Sheets.Refused);

            book.Sheets[0].Set("A1", "12");
            s.Check("once nothing reads it, it can",
                Sheets.Remove(book, book.Sheets.ToList().FindIndex(x => x.Name == "Detail")) is Sheets.Done);

            var back = Reopen(file);
            s.Check("and it is gone when the file opens again", !back.Workbook!.Sheets.Any(x => x.Name == "Detail"));
            s.Check("the last sheet cannot be taken out", Sheets.Remove(back.Workbook, 0) is Sheets.Refused);
        }

        // ---- where the contents sit, and what colour ----
        {
            var file = Fresh();
            var sheet = file.Workbook!.Sheets[0];
            sheet.Set("A1", "Pine Street");
            sheet.Set("B1", "1200");

            sheet.SetAlignment(new[] { CellRef.Parse("A1") }, "center", null);
            s.Check("a cell can be centred", sheet.AlignmentAt(CellRef.Parse("A1")).Horizontal == "center");
            sheet.SetAlignment(new[] { CellRef.Parse("A1") }, null, true);
            s.Check("and set to wrap", sheet.AlignmentAt(CellRef.Parse("A1")).Wrap);
            s.Check("without losing where it sits", sheet.AlignmentAt(CellRef.Parse("A1")).Horizontal == "center");
            sheet.SetAlignment(new[] { CellRef.Parse("A1") }, "general", null);
            s.Check("and can be put back to nothing in particular",
                sheet.AlignmentAt(CellRef.Parse("A1")).Horizontal == "general");

            sheet.SetColours(new[] { CellRef.Parse("B1") }, "FFE7A1", "7A4B00");
            s.Check("a cell can be given a colour behind it", sheet.ColoursAt(CellRef.Parse("B1")).Background == "FFE7A1");
            s.Check("and a colour for its words", sheet.ColoursAt(CellRef.Parse("B1")).Ink == "7A4B00");
            s.Check("a hash and an alpha pair are both accepted",
                Threw(() => sheet.SetColours(new[] { CellRef.Parse("B1") }, "#FFE7A1", "FF7A4B00")) == false);
            sheet.SetColours(new[] { CellRef.Parse("B1") }, "", "");
            s.Check("and both can be taken away", sheet.ColoursAt(CellRef.Parse("B1")).Background == "");

            sheet.SetColours(new[] { CellRef.Parse("B1") }, "C8E6C9", null);
            var back = Reopen(file);
            s.Check("colour survives a save", back.Workbook!.Sheets[0].ColoursAt(CellRef.Parse("B1")).Background == "C8E6C9");
            s.Check("and the workbook still opens", back.Workbook.Sheets.Count > 0);
        }

        // ---- how tall a row is ----
        {
            var file = Fresh();
            var sheet = file.Workbook!.Sheets[0];
            sheet.Set("A1", "tall");
            s.Check("a row says nothing about its height to begin with", sheet.HeightPoints(1) == 0);
            sheet.SetHeightPoints(1, 42);
            s.Check("a height set is a height read back", Math.Abs(sheet.HeightPoints(1) - 42) < 0.01);
            sheet.SetHeightPoints(1, 0);
            s.Check("and it can be put back to the sheet's own", sheet.HeightPoints(1) == 0);

            sheet.SetHeightPoints(2, 9000);
            s.Check("a silly height is brought back into range", sheet.HeightPoints(2) <= 409);
            sheet.SetHeightPoints(3, 30);
            var back = Reopen(file);
            s.Check("a height survives a save", Math.Abs(back.Workbook!.Sheets[0].HeightPoints(3) - 30) < 0.01);
        }

        // ---- what stays on screen ----
        {
            var file = Fresh();
            var sheet = file.Workbook!.Sheets[0];
            sheet.SetFrozen(2, 3);
            s.Check("columns can be held still", sheet.FrozenColumns == 2);
            s.Check("and rows at the same time", sheet.FrozenRows == 3);

            sheet.SetFrozenRows(1);
            s.Check("setting the rows alone leaves the columns held", sheet.FrozenColumns == 2);
            s.Check("and sets the rows", sheet.FrozenRows == 1);

            var back = Reopen(file);
            s.Check("both survive a save",
                back.Workbook!.Sheets[0].FrozenColumns == 2 && back.Workbook.Sheets[0].FrozenRows == 1);

            back.Workbook.Sheets[0].SetFrozen(0, 0);
            s.Check("and both can be let go", back.Workbook.Sheets[0].FrozenColumns == 0
                                           && back.Workbook.Sheets[0].FrozenRows == 0);
        }

        return s;
    }

    private static bool Threw(Action what)
    {
        try { what(); return false; } catch { return true; }
    }
}
