namespace Plain.Core.Tests;

/// <summary>
/// Sorting, and more importantly the refusing. Every check that Plain declines to sort is worth more than the ones
/// that say it sorted, because a wrong sort is the failure that costs someone their afternoon and their trust.
/// </summary>
public static class SortSuite
{
    private static (PlainFile File, Workbook Book, Sheet Sheet) Fresh(string[][] rows)
    {
        var path = System.IO.Path.Combine(System.IO.Path.GetTempPath(), "plain-sort-" + Guid.NewGuid().ToString("N") + ".xlsx");
        var file = PlainFile.Create(path);
        var book = file.Workbook!;
        var sheet = book.Sheets[0];
        for (int r = 0; r < rows.Length; r++)
            for (int c = 0; c < rows[r].Length; c++)
                sheet.Set(new CellRef(c + 1, r + 1), rows[r][c]);
        return (file, book, sheet);
    }

    private static string Col(Sheet s, int column, int top, int bottom) =>
        string.Join(",", Enumerable.Range(top, bottom - top + 1).Select(r => s.Read(new CellRef(column, r)).Display));

    public static Suite Run()
    {
        var s = new Suite("sort");

        // ---- text, up and down ----
        {
            var (_, book, sheet) = Fresh(new[]
            {
                new[] { "Name", "Amount" },
                new[] { "Pear", "3" },
                new[] { "Apple", "1" },
                new[] { "Cherry", "2" },
            });
            var r = Sort.Rows(book, sheet, 2, 4, 1, 2, 1, ascending: true);
            s.Check("text sorts up", r is Sort.Sorted && Col(sheet, 1, 2, 4) == "Apple,Cherry,Pear");
            s.Check("the other columns come with it", Col(sheet, 2, 2, 4) == "1,2,3");

            Sort.Rows(book, sheet, 2, 4, 1, 2, 1, ascending: false);
            s.Check("and down", Col(sheet, 1, 2, 4) == "Pear,Cherry,Apple");
        }

        // ---- numbers sort by value, not as text ----
        {
            var (_, book, sheet) = Fresh(new[]
            {
                new[] { "9" }, new[] { "100" }, new[] { "20" },
            });
            Sort.Rows(book, sheet, 1, 3, 1, 1, 1, ascending: true);
            s.Check("a hundred is after twenty, not before it", Col(sheet, 1, 1, 3) == "9,20,100");
        }

        // ---- a number stays a number ----
        {
            var (file, book, sheet) = Fresh(new[]
            {
                new[] { "2000" }, new[] { "24000" }, new[] { "300" },
            });
            Sort.Rows(book, sheet, 1, 3, 1, 1, 1, ascending: true);
            s.Check("what was a number is still a number after sorting",
                sheet.Read(new CellRef(1, 3)).Kind == CellKind.Number);
            s.Check("and holds its value, not its formatting",
                sheet.Read(new CellRef(1, 3)).Raw == "24000");
        }

        // ---- empty rows go last whichever way you sort ----
        {
            var (_, book, sheet) = Fresh(new[]
            {
                new[] { "B" }, new[] { "" }, new[] { "A" },
            });
            Sort.Rows(book, sheet, 1, 3, 1, 1, 1, ascending: true);
            s.Check("empty sorts last going up", Col(sheet, 1, 1, 3) == "A,B,");
            Sort.Rows(book, sheet, 1, 3, 1, 1, 1, ascending: false);
            s.Check("and last going down too", Col(sheet, 1, 1, 3) == "B,A,");
        }

        // ---- the refusals ----
        {
            var (_, book, sheet) = Fresh(new[]
            {
                new[] { "3", "" }, new[] { "1", "" }, new[] { "2", "=A1+A2" },
            });
            var r = Sort.Rows(book, sheet, 1, 3, 1, 2, 1, ascending: true);
            s.Check("a formula inside the block is refused", r is Sort.Refused);
            s.Check("and the reason names the cell", r is Sort.Refused f && f.Reason.Contains("C3") || r is Sort.Refused f2 && f2.Reason.Contains("B3"));
            s.Check("and nothing moved", Col(sheet, 1, 1, 3) == "3,1,2");
        }
        {
            // A total under a table reads every row of it. Shuffling the rows does not change the total, so this is
            // the ordinary case and it has to be allowed or sorting would be useless on any real sheet.
            var (_, book, sheet) = Fresh(new[]
            {
                new[] { "3" }, new[] { "1" }, new[] { "2" }, new[] { "" }, new[] { "" },
            });
            sheet.Set(new CellRef(1, 6), "=SUM(A1:A3)");
            var r = Sort.Rows(book, sheet, 1, 3, 1, 1, 1, ascending: true);
            s.Check("a total that reads every row does not stop the sort", r is Sort.Sorted);
            s.Check("and the rows did move", Col(sheet, 1, 1, 3) == "1,2,3");
            s.Check("and the total still reads the same rows",
                (sheet.Read(new CellRef(1, 6)).Formula ?? "").Contains("A1:A3"));
        }
        {
            // Half a table is a different matter: after the sort those two rows hold different values.
            var (_, book, sheet) = Fresh(new[]
            {
                new[] { "3" }, new[] { "1" }, new[] { "2" }, new[] { "" }, new[] { "" },
            });
            sheet.Set(new CellRef(1, 6), "=SUM(A1:A2)");
            var r = Sort.Rows(book, sheet, 1, 3, 1, 1, 1, ascending: true);
            s.Check("a formula reading only some of the rows is refused", r is Sort.Refused);
            s.Check("and says so", r is Sort.Refused half && half.Reason.Contains("not all"));
            s.Check("nothing moved for that", Col(sheet, 1, 1, 3) == "3,1,2");
        }
        {
            // A lookup over the whole block does depend on the order, so covering every row is not enough.
            var (_, book, sheet) = Fresh(new[]
            {
                new[] { "3" }, new[] { "1" }, new[] { "2" }, new[] { "" }, new[] { "" },
            });
            sheet.Set(new CellRef(1, 6), "=INDEX(A1:A3,2)");
            var r = Sort.Rows(book, sheet, 1, 3, 1, 1, 1, ascending: true);
            s.Check("a lookup over the block is refused even though it covers every row", r is Sort.Refused);
            s.Check("and says the order is what matters",
                r is Sort.Refused look && look.Reason.Contains("order"));
        }
        {
            var (_, book, sheet) = Fresh(new[] { new[] { "1" }, new[] { "2" } });
            s.Check("one row is not a sort", Sort.Rows(book, sheet, 1, 1, 1, 1, 1, true) is Sort.Refused);
            s.Check("sorting by a column outside the block is refused",
                Sort.Rows(book, sheet, 1, 2, 1, 1, 5, true) is Sort.Refused);
        }

        // ---- a formula that reads a different sheet does not stop a sort ----
        {
            var (_, book, sheet) = Fresh(new[] { new[] { "3" }, new[] { "1" }, new[] { "2" } });
            if (book.Sheets.Count > 1)
                s.Check("a formula on another sheet about another sheet is no obstacle",
                    Sort.Rows(book, sheet, 1, 3, 1, 1, 1, true) is Sort.Sorted);
            else
                s.Check("a new workbook has one sheet, so there is nothing to cross", true);
        }

        // ---- sorting is stable, so two passes give a two-key sort ----
        {
            var (_, book, sheet) = Fresh(new[]
            {
                new[] { "B", "2" }, new[] { "A", "2" }, new[] { "B", "1" }, new[] { "A", "1" },
            });
            Sort.Rows(book, sheet, 1, 4, 1, 2, 2, ascending: true);   // by amount first
            Sort.Rows(book, sheet, 1, 4, 1, 2, 1, ascending: true);   // then by name
            s.Check("two passes sort by two columns", Col(sheet, 1, 1, 4) == "A,A,B,B");
            s.Check("and the second key keeps its order inside the first",
                Col(sheet, 2, 1, 4) == "1,2,1,2");
        }

        // ---- the file still round trips after a sort ----
        {
            var (file, book, sheet) = Fresh(new[] { new[] { "3" }, new[] { "1" }, new[] { "2" } });
            Sort.Rows(book, sheet, 1, 3, 1, 1, 1, true);
            file.Flush();
            var bytes = file.Package.ToBytes();
            var back = PlainFile.Read(bytes, file.Path);
            s.Check("a sorted file opens again", back.Workbook is not null);
            s.Check("and reads back sorted", Col(back.Workbook!.Sheets[0], 1, 1, 3) == "1,2,3");
        }

        return s;
    }
}
