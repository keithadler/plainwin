namespace Plain.Core.Tests;

public static class FormulaSuite
{
    /// <summary>A little sheet to read from: A1=10, A2=20, A3=30, B1="hi", B2 blank, C1=TRUE.</summary>
    private static Value Cells(string? sheet, CellRef cell)
    {
        if (sheet is not null && sheet != "Data") return Value.Blank;
        return (cell.Column, cell.Row) switch
        {
            (1, 1) => Value.Of(10), (1, 2) => Value.Of(20), (1, 3) => Value.Of(30), (1, 4) => Value.Of(-5),
            (2, 1) => Value.Of("hi"),
            (3, 1) => Value.Of(true),
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

        // The whole point: anything not understood must be refused, never guessed.
        s.Equal("an unknown function is refused", "(not understood)", Eval("=VLOOKUP(A1,A1:B3,2,FALSE)"));
        s.Equal("a defined name is refused", "(not understood)", Eval("=TaxRate*A1"));
        s.Equal("today is refused, because it would not stay true", "(not understood)", Eval("=TODAY()"));
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
