namespace Plain.Core.Tests;

public static class RefsSuite
{
    private static string All(string formula) =>
        string.Join(' ', Refs.Parse(formula).Select(r => r.ToString()));

    public static Suite Run()
    {
        var s = new Suite("refs");

        s.Equal("a single cell", "A1", All("=A1"));
        s.Equal("two cells added", "A1 B2", All("=A1+B2"));
        s.Equal("a range inside a function", "B2:B4", All("=SUM(B2:B4)"));
        s.Equal("the function name is not a reference", "A1", All("=LOG10(A1)"));
        s.Equal("a function whose name ends in digits", "A1", All("=T1(A1)"));
        s.Equal("dollars are ignored", "A1", All("=$A$1"));
        s.Equal("mixed dollars", "B7", All("=B$7"));
        s.Equal("several arguments", "A1 B1 C1", All("=IF(A1>0,B1,C1)"));
        s.Equal("a range across columns and rows", "A1:C9", All("=SUM(A1:C9)"));
        s.Equal("a backwards range is normalised", "A1:C9", All("=SUM(C9:A1)"));

        s.Equal("a sheet qualified cell", "Sheet2!A1", All("=Sheet2!A1"));
        s.Equal("a sheet qualified range", "Sheet2!A1:B2", All("=SUM(Sheet2!A1:B2)"));
        s.Equal("a quoted sheet name", "My Sheet!A1:B2", All("='My Sheet'!A1:B2"));
        s.Check("an escaped quote in a sheet name", Refs.Parse("='It''s'!A1").Single().Sheet == "It's");

        s.Equal("text in quotes is not a reference", "", All("=\"A1\""));
        s.Equal("text beside a reference", "A1", All("=A1&\" B2\""));
        s.Equal("a bare name is not a reference", "", All("=MyRange"));
        s.Equal("TRUE is not a column", "", All("=TRUE"));
        s.Equal("a table reference is left alone", "", All("=Table1[Amount]"));
        s.Equal("a row past the last one is not a reference", "", All("=A1048577"));
        s.Equal("a column past the last one is not a reference", "", All("=AAAA1"));
        s.Equal("the last real cell is a reference", "XFD1048576", All("=XFD1048576"));
        s.Equal("a number on its own is not a cell", "", All("=42+7"));

        var wholeColumn = Refs.Parse("=SUM(A:A)");
        s.Equal("a whole column is one range", 1, wholeColumn.Count);
        s.Check("a whole column covers every row", wholeColumn[0].Contains(new CellRef(1, 500000)));
        s.Check("a whole column covers only that column", !wholeColumn[0].Contains(new CellRef(2, 1)));

        var wholeRow = Refs.Parse("=SUM(3:3)");
        s.Equal("a whole row is one range", 1, wholeRow.Count);
        s.Check("a whole row covers every column", wholeRow[0].Contains(new CellRef(9000, 3)));
        s.Check("a whole row covers only that row", !wholeRow[0].Contains(new CellRef(1, 4)));

        var range = Refs.Parse("=SUM(B2:D5)").Single();
        s.Check("a range holds its corners", range.Contains(new CellRef(2, 2)) && range.Contains(new CellRef(4, 5)));
        s.Check("a range holds its middle", range.Contains(new CellRef(3, 3)));
        s.Check("a range stops at its edges", !range.Contains(new CellRef(1, 2)) && !range.Contains(new CellRef(2, 6)));

        // Nothing may send the scanner round for ever or off the end.
        foreach (var odd in new[] { "=", "=(((", "='unclosed", "=\"unclosed", "=A", "=!", "=Sheet2!", "=:", "=$", "=A1:", "==A1" })
            s.Check($"awkward input is survived: {odd}", Refs.Parse(odd) is not null);

        // How a block of cells travels on the clipboard.
        s.Equal("one cell writes as itself", "a", Tabular.Write(new[] { new[] { "a" } }));
        s.Equal("a row is tab separated", "a\tb", Tabular.Write(new[] { new[] { "a", "b" } }));
        s.Equal("rows are separated the way Windows does", "a\r\nb", Tabular.Write(new[] { new[] { "a" }, new[] { "b" } }));

        s.Equal("one line reads as one row", 1, Tabular.Read("a\tb").Count);
        s.Equal("and as two cells", 2, Tabular.Read("a\tb")[0].Count);
        s.Equal("windows line endings", 2, Tabular.Read("a\r\nb").Count);
        s.Equal("unix line endings", 2, Tabular.Read("a\nb").Count);
        s.Equal("old mac line endings", 2, Tabular.Read("a\rb").Count);
        s.Equal("a trailing newline does not add a row", 2, Tabular.Read("a\r\nb\r\n").Count);
        s.Equal("two trailing newlines do add an empty one", 3, Tabular.Read("a\nb\n\n").Count);
        s.Equal("empty text is no rows", 0, Tabular.Read("").Count);
        s.Equal("an empty cell survives between tabs", "", Tabular.Read("a\t\tb")[0][1]);
        s.Equal("ragged rows keep their own widths", 3, Tabular.Width(Tabular.Read("a\nb\tc\td")));
        s.Check("a block survives a round trip",
                Tabular.Read(Tabular.Write(new[] { new[] { "1", "two" }, new[] { "", "=SUM(A1:A2)" } }))
                       .SelectMany(r => r).SequenceEqual(new[] { "1", "two", "", "=SUM(A1:A2)" }));

        return s;
    }
}
