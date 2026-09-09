namespace Plain.Core.Tests;

public static class FormulaSuite
{
    /// <summary>
    /// A little sheet to read from. Column A holds numbers, B some text, C a yes/no. Columns E to G hold a small
    /// table to look things up in: E is a name, F a region, G an amount.
    /// </summary>
    private static Value Cells(string? sheet, CellRef cell)
    {
        if (sheet is not null && sheet != "Data") return Value.Blank;
        return (cell.Column, cell.Row) switch
        {
            (1, 1) => Value.Of(10), (1, 2) => Value.Of(20), (1, 3) => Value.Of(30), (1, 4) => Value.Of(-5),
            (2, 1) => Value.Of("hi"),
            (3, 1) => Value.Of(true),

            (5, 1) => Value.Of("apples"),  (6, 1) => Value.Of("north"), (7, 1) => Value.Of(100),
            (5, 2) => Value.Of("pears"),   (6, 2) => Value.Of("south"), (7, 2) => Value.Of(250),
            (5, 3) => Value.Of("plums"),   (6, 3) => Value.Of("north"), (7, 3) => Value.Of(75),
            (5, 4) => Value.Of("quinces"), (6, 4) => Value.Of("south"), (7, 4) => Value.Of(400),

            (9, 1) => Value.Of(45000),   // a date serial, 2023-03-15
            _ => Value.Blank,
        };
    }

    private static string Eval(string formula)
    {
        if (!Formula.TryEvaluate(formula, Cells, out var value)) return "(not understood)";
        return value.Kind switch
        {
            Value.Sort.Number => value.Number.ToString("0.############", System.Globalization.CultureInfo.InvariantCulture),
            Value.Sort.Bool => value.Number != 0 ? "TRUE" : "FALSE",
            Value.Sort.Blank => "",
            _ => value.Text,
        };
    }

    public static Suite Run()
    {
        var s = new Suite("formula");

        s.Equal("a number", "42", Eval("=42"));
        s.Equal("adding", "3", Eval("=1+2"));
        s.Equal("order of operations", "7", Eval("=1+2*3"));
        s.Equal("brackets first", "9", Eval("=(1+2)*3"));
        s.Equal("subtracting", "-1", Eval("=1-2"));
        s.Equal("dividing", "2.5", Eval("=5/2"));
        s.Equal("powers", "8", Eval("=2^3"));
        s.Equal("a minus in front", "-3", Eval("=-3"));
        s.Equal("minus binds tighter than plus", "1", Eval("=-2+3"));
        s.Equal("a percentage", "0.05", Eval("=5%"));
        s.Equal("decimals", "0.75", Eval("=0.75"));
        s.Equal("scientific notation", "1500", Eval("=1.5e3"));
        s.Equal("dividing by nothing is an error", "#DIV/0!", Eval("=1/0"));

        s.Equal("a cell", "10", Eval("=A1"));
        s.Equal("two cells", "30", Eval("=A1+A2"));
        s.Equal("a blank cell counts as nothing", "10", Eval("=A1+B2"));
        s.Equal("a cell on a named sheet", "20", Eval("=Data!A2"));
        s.Equal("a cell on a sheet with no values", "0", Eval("=Elsewhere!A2+0"));

        s.Equal("a total", "60", Eval("=SUM(A1:A3)"));
        s.Equal("a total of scattered things", "40", Eval("=SUM(A1,A3)"));
        s.Equal("a total mixing a range and a number", "61", Eval("=SUM(A1:A3,1)"));
        s.Equal("an average", "20", Eval("=AVERAGE(A1:A3)"));
        s.Equal("the smallest", "-5", Eval("=MIN(A1:A4)"));
        s.Equal("the largest", "30", Eval("=MAX(A1:A3)"));
        s.Equal("counting numbers", "4", Eval("=COUNT(A1:A9)"));
        s.Equal("counting anything", "4", Eval("=COUNTA(A1:A9)"));
        s.Equal("text is not counted as a number", "0", Eval("=COUNT(B1)"));
        s.Equal("rounding", "3.14", Eval("=ROUND(3.14159,2)"));
        s.Equal("rounding half away from zero", "3", Eval("=ROUND(2.5,0)"));
        s.Equal("absolute", "5", Eval("=ABS(-5)"));
        s.Equal("square root", "4", Eval("=SQRT(16)"));
        s.Equal("the root of a negative is an error", "#NUM!", Eval("=SQRT(-1)"));
        s.Equal("remainder", "1", Eval("=MOD(7,3)"));
        s.Equal("whole part", "3", Eval("=INT(3.9)"));

        s.Equal("a test that passes", "yes", Eval("=IF(A1>5,\"yes\",\"no\")"));
        s.Equal("a test that fails", "no", Eval("=IF(A1>50,\"yes\",\"no\")"));
        s.Equal("a test on equality", "TRUE", Eval("=A1=10"));
        s.Equal("not equal", "TRUE", Eval("=A1<>11"));
        s.Equal("less than or equal", "TRUE", Eval("=A1<=10"));
        s.Equal("and", "TRUE", Eval("=AND(A1=10,A2=20)"));
        s.Equal("or", "TRUE", Eval("=OR(A1=99,A2=20)"));
        s.Equal("not", "TRUE", Eval("=NOT(A1=99)"));
        s.Equal("catching an error", "safe", Eval("=IFERROR(1/0,\"safe\")"));
        s.Equal("an error travels up", "#DIV/0!", Eval("=SUM(1/0,2)"));

        s.Equal("text", "hi", Eval("=B1"));
        s.Equal("joining text", "hi there", Eval("=B1&\" there\""));
        s.Equal("length", "2", Eval("=LEN(B1)"));
        s.Equal("upper", "HI", Eval("=UPPER(B1)"));
        s.Equal("the left of some text", "h", Eval("=LEFT(B1,1)"));
        s.Equal("the right of some text", "i", Eval("=RIGHT(B1,1)"));
        s.Equal("the middle", "ell", Eval("=MID(\"hello\",2,3)"));
        s.Equal("trimming", "a b", Eval("=TRIM(\"  a b  \")"));
        s.Equal("a quote inside text", "say \"hi\"", Eval("=\"say \"\"hi\"\"\""));
        s.Equal("joining a number to text", "n10", Eval("=\"n\"&A1"));

        s.Equal("spaces are ignored", "30", Eval("= A1 + A2 "));
        s.Equal("a formula with no equals sign", "3", Eval("1+2"));

        // Looking things up in a table.
        s.Equal("a lookup finds the row and takes a column", "250", Eval("=VLOOKUP(\"pears\",E1:G4,3,FALSE)"));
        s.Equal("a lookup is not case sensitive", "250", Eval("=VLOOKUP(\"PEARS\",E1:G4,3,FALSE)"));
        s.Equal("a lookup that finds nothing says so", "#N/A", Eval("=VLOOKUP(\"kiwi\",E1:G4,3,FALSE)"));
        s.Equal("a lookup can take the middle column", "south", Eval("=VLOOKUP(\"quinces\",E1:G4,2,FALSE)"));
        s.Equal("across instead of down", "south", Eval("=HLOOKUP(\"north\",E1:G2,2,FALSE)"));
        s.Equal("the newer lookup", "75", Eval("=XLOOKUP(\"plums\",E1:E4,G1:G4)"));
        s.Equal("with something to say when it fails", "none", Eval("=XLOOKUP(\"kiwi\",E1:E4,G1:G4,\"none\")"));
        s.Equal("finding the position of something", "2", Eval("=MATCH(\"pears\",E1:E4,0)"));
        s.Equal("taking a cell out of a table by position", "400", Eval("=INDEX(G1:G4,4)"));
        s.Equal("by row and column", "south", Eval("=INDEX(E1:G4,2,2)"));
        s.Equal("position and index together", "75", Eval("=INDEX(G1:G4,MATCH(\"plums\",E1:E4,0))"));

        // Totalling with a condition.
        s.Equal("adding up what matches", "175", Eval("=SUMIF(F1:F4,\"north\",G1:G4)"));
        s.Equal("counting what matches", "2", Eval("=COUNTIF(F1:F4,\"south\")"));
        s.Equal("a comparison as the condition", "650", Eval("=SUMIF(G1:G4,\">100\")"));
        s.Equal("counting with a comparison", "2", Eval("=COUNTIF(G1:G4,\"<=100\")"));
        s.Equal("a wildcard condition", "2", Eval("=COUNTIF(E1:E4,\"p*\")"));
        s.Equal("the average of what matches", "87.5", Eval("=AVERAGEIF(F1:F4,\"north\",G1:G4)"));
        s.Equal("two conditions at once", "400", Eval("=SUMIFS(G1:G4,F1:F4,\"south\",G1:G4,\">300\")"));
        s.Equal("counting on two conditions", "1", Eval("=COUNTIFS(F1:F4,\"north\",G1:G4,\">80\")"));
        s.Equal("a condition that matches nothing totals nothing", "0", Eval("=SUMIF(F1:F4,\"east\",G1:G4)"));

        // Dates.
        s.Equal("the year of a date", "2023", Eval("=YEAR(I1)"));
        s.Equal("the month", "3", Eval("=MONTH(I1)"));
        s.Equal("the day", "15", Eval("=DAY(I1)"));
        s.Equal("building a date", "45000", Eval("=DATE(2023,3,15)"));
        s.Equal("the end of the month", "45016", Eval("=EOMONTH(I1,0)"));
        s.Equal("a month later", "45031", Eval("=EDATE(I1,1)"));
        s.Equal("days between", "31", Eval("=DAYS(DATE(2023,4,15),I1)"));
        s.Check("today is a date in this century", double.Parse(Eval("=TODAY()")) > 44000);

        // The whole point: anything not understood must be refused, never guessed.
        s.Equal("an unknown function is refused", "(not understood)", Eval("=BAHTTEXT(A1)"));
        s.Equal("a defined name is refused", "(not understood)", Eval("=TaxRate*A1"));
        s.Equal("rubbish is refused", "(not understood)", Eval("=1+"));
        s.Equal("an unclosed bracket is refused", "(not understood)", Eval("=SUM(A1:A3"));
        s.Equal("an unclosed quote is refused", "(not understood)", Eval("=\"abc"));
        s.Equal("trailing rubbish is refused", "(not understood)", Eval("=1+2 xyzzy"));
        s.Equal("an empty formula is refused", "(not understood)", Eval("="));
        s.Equal("a reference to another file is refused", "(not understood)", Eval("=[Book2]Sheet1!A1"));

        // An error written into the formula is carried, not invented.
        s.Equal("a reference error is kept", "#REF!", Eval("=#REF!"));
        s.Equal("and travels", "#REF!", Eval("=#REF!+1"));

        // Nothing may loop for ever or run away with the machine.
        s.Equal("a whole column is refused rather than walked", "(not understood)", Eval("=SUM(A:A)"));
        foreach (var odd in new[] { "=(((((((", "=,,,", "=)", "=\"\"\"", "=A1:", "=SUM(,)", "=--", "=1e" })
            s.Check($"awkward input is survived: {odd}", Eval(odd) is not null);

        return s;
    }
}
