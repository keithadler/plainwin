using System.Globalization;

namespace Plain.Core;

/// <summary>What a formula came to: a number, some text, a yes/no, or an error.</summary>
public readonly record struct Value
{
    public enum Sort { Number, Text, Bool, Error, Blank }

    public Sort Kind { get; private init; }
    public double Number { get; private init; }
    public string Text { get; private init; }

    public static Value Of(double number) => new() { Kind = Sort.Number, Number = number, Text = "" };
    public static Value Of(string text) => new() { Kind = Sort.Text, Number = 0, Text = text };
    public static Value Of(bool yes) => new() { Kind = Sort.Bool, Number = yes ? 1 : 0, Text = "" };
    public static Value Error(string code) => new() { Kind = Sort.Error, Number = 0, Text = code };
    public static readonly Value Blank = new() { Kind = Sort.Blank, Number = 0, Text = "" };

    public bool IsError => Kind == Sort.Error;

    /// <summary>The number this stands for, or an error when it does not stand for one.</summary>
    public bool AsNumber(out double number)
    {
        switch (Kind)
        {
            case Sort.Number or Sort.Bool: number = Number; return true;
            case Sort.Blank: number = 0; return true;
            case Sort.Text:
                return double.TryParse(Text, NumberStyles.Float, CultureInfo.InvariantCulture, out number);
            default: number = 0; return false;
        }
    }

    public string AsText() => Kind switch
    {
        Sort.Number => Number.ToString("R", CultureInfo.InvariantCulture),
        Sort.Bool => Number != 0 ? "TRUE" : "FALSE",
        Sort.Blank => "",
        _ => Text,
    };
}

/// <summary>
/// Works out what a formula comes to, but only when it understands every part of it. Anything unfamiliar - a
/// function not on the list, a name, a reference to another workbook - and it gives up and says so, because a
/// spreadsheet that shows a confidently wrong total is worse than one that shows nothing.
///
/// Excel is still asked to recalculate the file when it opens, so this is what you see while you work rather than
/// the last word on what the number is.
/// </summary>
public sealed class Formula
{
    /// <summary>Where cell values come from. Returns Blank for an empty cell.</summary>
    public delegate Value Lookup(string? sheet, CellRef cell);

    private readonly Lookup _lookup;
    private readonly string _text;
    private int _at;

    private Formula(string text, Lookup lookup) { _text = text; _lookup = lookup; }

    /// <summary>Work out a formula. Returns false when any part of it was not understood.</summary>
    public static bool TryEvaluate(string formula, Lookup lookup, out Value value)
    {
        value = Value.Blank;
        if (string.IsNullOrWhiteSpace(formula)) return false;
        var text = formula.StartsWith('=') ? formula[1..] : formula;
        var parser = new Formula(text, lookup);
        try
        {
            var result = parser.Expression();
            parser.SkipSpace();
            if (parser._at != parser._text.Length) return false;   // trailing rubbish: do not guess
            value = parser.Flatten(result);
            return true;
        }
        catch (NotSupportedException) { return false; }
        catch (Exception) { return false; }
    }

    // ---------- an operand is a value, or a range waiting to be flattened ----------

    private readonly record struct Operand(Value Value, RefRange? Range)
    {
        public static Operand Of(Value v) => new(v, null);
        public static Operand Of(RefRange r) => new(Value.Blank, r);
    }

    private Value Flatten(Operand operand)
    {
        if (operand.Range is not { } range) return operand.Value;
        // A range used where one value is wanted is only meaningful when it holds exactly one cell.
        if (range.ColumnMin != range.ColumnMax || range.RowMin != range.RowMax) throw new NotSupportedException();
        return _lookup(range.Sheet, new CellRef(range.ColumnMin, range.RowMin));
    }

    /// <summary>One cell of a range, counted from its top left corner.</summary>
    private Value At(RefRange range, int rowOffset, int columnOffset)
    {
        int r = range.RowMin + rowOffset, c = range.ColumnMin + columnOffset;
        if (r > range.RowMax || c > range.ColumnMax || rowOffset < 0 || columnOffset < 0) return Value.Error("#REF!");
        return _lookup(range.Sheet, new CellRef(c, r));
    }

    private static int Rows(RefRange range) => range.RowMax - range.RowMin + 1;
    private static int Columns(RefRange range) => range.ColumnMax - range.ColumnMin + 1;

    private RefRange RangeOf(Operand operand) =>
        operand.Range ?? throw new NotSupportedException();   // a function wanting a range was handed a single value

    /// <summary>
    /// Does a value satisfy a criterion of the kind SUMIF and COUNTIF take: a number, a comparison like "&gt;100",
    /// or text with * and ? standing for any run and any one character.
    /// </summary>
    private static bool Meets(Value value, Value criterion)
    {
        string text = criterion.AsText().Trim();
        string op = "";
        foreach (var candidate in new[] { "<>", "<=", ">=", "<", ">", "=" })
            if (text.StartsWith(candidate, StringComparison.Ordinal)) { op = candidate; text = text[candidate.Length..].Trim(); break; }

        if (double.TryParse(text, NumberStyles.Float, CultureInfo.InvariantCulture, out var wanted)
            && value.AsNumber(out var have) && value.Kind is not Value.Sort.Text)
        {
            return op switch
            {
                "<" => have < wanted, "<=" => have <= wanted,
                ">" => have > wanted, ">=" => have >= wanted,
                "<>" => have != wanted, _ => have == wanted,
            };
        }

        string actual = value.AsText();
        bool same = Like(actual, text);
        return op == "<>" ? !same : same;
    }

    /// <summary>Text matching with * and ?, the way a spreadsheet criterion means them.</summary>
    private static bool Like(string text, string pattern)
    {
        if (!pattern.Contains('*') && !pattern.Contains('?'))
            return string.Equals(text, pattern, StringComparison.CurrentCultureIgnoreCase);

        var built = new System.Text.StringBuilder("^");
        foreach (char c in pattern)
            built.Append(c switch { '*' => ".*", '?' => ".", _ => System.Text.RegularExpressions.Regex.Escape(c.ToString()) });
        built.Append('$');
        return System.Text.RegularExpressions.Regex.IsMatch(text, built.ToString(),
            System.Text.RegularExpressions.RegexOptions.IgnoreCase | System.Text.RegularExpressions.RegexOptions.CultureInvariant,
            TimeSpan.FromMilliseconds(200));
    }

    private static readonly DateTime Epoch = new(1899, 12, 30);
    private static double Serial(DateTime when) => (when.Date - Epoch).TotalDays;
    private static DateTime FromSerial(double serial) => Epoch.AddDays(serial);

    private IEnumerable<Value> Spread(Operand operand)
    {
        if (operand.Range is not { } range) { yield return operand.Value; yield break; }
        long width = (long)range.ColumnMax - range.ColumnMin + 1;
        long height = (long)range.RowMax - range.RowMin + 1;
        if (width * height > 1_000_000) throw new NotSupportedException();   // a whole-column sum of nothing useful
        for (int c = range.ColumnMin; c <= range.ColumnMax; c++)
            for (int r = range.RowMin; r <= range.RowMax; r++)
                yield return _lookup(range.Sheet, new CellRef(c, r));
    }

    // ---------- the grammar ----------

    private void SkipSpace() { while (_at < _text.Length && _text[_at] is ' ' or '\t' or '\n' or '\r') _at++; }

    private bool Take(string token)
    {
        SkipSpace();
        if (_at + token.Length > _text.Length) return false;
        if (!_text.AsSpan(_at, token.Length).SequenceEqual(token)) return false;
        _at += token.Length;
        return true;
    }

    private Operand Expression()
    {
        var left = Concat();
        while (true)
        {
            SkipSpace();
            string? op = Take("<>") ? "<>" : Take("<=") ? "<=" : Take(">=") ? ">=" :
                         Take("=") ? "=" : Take("<") ? "<" : Take(">") ? ">" : null;
            if (op is null) return left;
            var right = Concat();
            left = Operand.Of(Compare(Flatten(left), Flatten(right), op));
        }
    }

    private Operand Concat()
    {
        var left = Additive();
        while (Take("&"))
        {
            var right = Additive();
            var a = Flatten(left); var b = Flatten(right);
            if (a.IsError) return Operand.Of(a);
            if (b.IsError) return Operand.Of(b);
            left = Operand.Of(Value.Of(a.AsText() + b.AsText()));
        }
        return left;
    }

    private Operand Additive()
    {
        var left = Multiplicative();
        while (true)
        {
            SkipSpace();
            bool plus = Take("+");
            bool minus = !plus && Take("-");
            if (!plus && !minus) return left;
            var right = Multiplicative();
            left = Operand.Of(Arithmetic(Flatten(left), Flatten(right), plus ? '+' : '-'));
        }
    }

    private Operand Multiplicative()
    {
        var left = Unary();
        while (true)
        {
            SkipSpace();
            bool times = Take("*");
            bool over = !times && Take("/");
            if (!times && !over) return left;
            var right = Unary();
            left = Operand.Of(Arithmetic(Flatten(left), Flatten(right), times ? '*' : '/'));
        }
    }

    private Operand Unary()
    {
        SkipSpace();
        if (Take("-"))
        {
            var inner = Flatten(Unary());
            if (inner.IsError) return Operand.Of(inner);
            if (!inner.AsNumber(out var n)) return Operand.Of(Value.Error("#VALUE!"));
            return Operand.Of(Value.Of(-n));
        }
        if (Take("+")) return Unary();
        return Power();
    }

    private Operand Power()
    {
        var left = Postfix();
        SkipSpace();
        if (!Take("^")) return left;
        var right = Flatten(Unary());
        var baseValue = Flatten(left);
        if (baseValue.IsError) return Operand.Of(baseValue);
        if (right.IsError) return Operand.Of(right);
        if (!baseValue.AsNumber(out var a) || !right.AsNumber(out var b)) return Operand.Of(Value.Error("#VALUE!"));
        return Operand.Of(Value.Of(Math.Pow(a, b)));
    }

    private Operand Postfix()
    {
        var value = Primary();
        SkipSpace();
        while (_at < _text.Length && _text[_at] == '%')
        {
            _at++;
            var inner = Flatten(value);
            if (!inner.AsNumber(out var n)) return Operand.Of(Value.Error("#VALUE!"));
            value = Operand.Of(Value.Of(n / 100));
            SkipSpace();
        }
        return value;
    }

    private Operand Primary()
    {
        SkipSpace();
        if (_at >= _text.Length) throw new NotSupportedException();

        char ch = _text[_at];

        if (ch == '(')
        {
            _at++;
            var inner = Expression();
            if (!Take(")")) throw new NotSupportedException();
            return inner;
        }

        if (ch == '"')
        {
            _at++;
            var built = new System.Text.StringBuilder();
            while (_at < _text.Length)
            {
                if (_text[_at] == '"')
                {
                    if (_at + 1 < _text.Length && _text[_at + 1] == '"') { built.Append('"'); _at += 2; continue; }
                    _at++;
                    return Operand.Of(Value.Of(built.ToString()));
                }
                built.Append(_text[_at++]);
            }
            throw new NotSupportedException();
        }

        if (char.IsAsciiDigit(ch) || (ch == '.' && _at + 1 < _text.Length && char.IsAsciiDigit(_text[_at + 1])))
        {
            int start = _at;
            while (_at < _text.Length && (char.IsAsciiDigit(_text[_at]) || _text[_at] == '.')) _at++;
            if (_at < _text.Length && (_text[_at] is 'e' or 'E'))
            {
                int save = _at;
                _at++;
                if (_at < _text.Length && (_text[_at] is '+' or '-')) _at++;
                if (_at < _text.Length && char.IsAsciiDigit(_text[_at])) { while (_at < _text.Length && char.IsAsciiDigit(_text[_at])) _at++; }
                else _at = save;
            }
            var span = _text[start.._at];
            if (!double.TryParse(span, NumberStyles.Float, CultureInfo.InvariantCulture, out var number)) throw new NotSupportedException();
            return Operand.Of(Value.Of(number));
        }

        if (ch == '#') return Operand.Of(Value.Error(ReadErrorLiteral()));

        // A name: a function call, a reference, TRUE/FALSE, or something Plain will not guess at.
        return NameOrReference();
    }

    private string ReadErrorLiteral()
    {
        foreach (var code in new[] { "#REF!", "#VALUE!", "#DIV/0!", "#NAME?", "#N/A", "#NULL!", "#NUM!" })
            if (Take(code)) return code;
        throw new NotSupportedException();
    }

    private Operand NameOrReference()
    {
        int start = _at;

        // A reference, possibly naming a sheet, is the common case: let the reference scanner decide.
        var tokens = Refs.Scan(_text[_at..]);
        if (tokens.Count > 0 && tokens[0].Start == 0)
        {
            var token = tokens[0];
            _at += token.Length;
            SkipSpace();
            // A name followed by "(" was a function, and the scanner would have refused it, so this really is a reference.
            return Operand.Of(token.Range);
        }

        while (_at < _text.Length && (char.IsAsciiLetterOrDigit(_text[_at]) || _text[_at] is '_' or '.')) _at++;
        if (_at == start) throw new NotSupportedException();
        string name = _text[start.._at].ToUpperInvariant();

        if (name == "TRUE") return Operand.Of(Value.Of(true));
        if (name == "FALSE") return Operand.Of(Value.Of(false));

        SkipSpace();
        if (_at >= _text.Length || _text[_at] != '(') throw new NotSupportedException();   // a defined name: not guessed at
        _at++;

        var arguments = new List<Operand>();
        SkipSpace();
        if (_at < _text.Length && _text[_at] == ')') _at++;
        else
        {
            while (true)
            {
                arguments.Add(Expression());
                SkipSpace();
                if (Take(",")) continue;
                if (Take(")")) break;
                throw new NotSupportedException();
            }
        }
        return Operand.Of(Call(name, arguments));
    }

    // ---------- the functions Plain knows ----------

    private Value Call(string name, List<Operand> args)
    {
        IEnumerable<Value> All() => args.SelectMany(Spread);
        List<double> Numbers()
        {
            var list = new List<double>();
            foreach (var v in All())
            {
                if (v.IsError) throw new ErrorValue(v);
                if (v.Kind is Value.Sort.Number or Value.Sort.Bool) list.Add(v.Number);
            }
            return list;
        }
        double One(int i)
        {
            var v = Flatten(args[i]);
            if (v.IsError) throw new ErrorValue(v);
            if (!v.AsNumber(out var n)) throw new ErrorValue(Value.Error("#VALUE!"));
            return n;
        }
        string Str(int i)
        {
            var v = Flatten(args[i]);
            if (v.IsError) throw new ErrorValue(v);
            return v.AsText();
        }

        try
        {
            switch (name)
            {
                case "SUM": return Value.Of(Numbers().Sum());
                case "PRODUCT": return Value.Of(Numbers().Aggregate(1.0, (a, b) => a * b));
                case "AVERAGE":
                {
                    var numbers = Numbers();
                    return numbers.Count == 0 ? Value.Error("#DIV/0!") : Value.Of(numbers.Average());
                }
                case "MIN": { var n = Numbers(); return Value.Of(n.Count == 0 ? 0 : n.Min()); }
                case "MAX": { var n = Numbers(); return Value.Of(n.Count == 0 ? 0 : n.Max()); }
                case "COUNT": return Value.Of(All().Count(v => v.Kind == Value.Sort.Number));
                case "COUNTA": return Value.Of(All().Count(v => v.Kind != Value.Sort.Blank));
                case "COUNTBLANK": return Value.Of(All().Count(v => v.Kind == Value.Sort.Blank));
                case "ABS": return Value.Of(Math.Abs(One(0)));
                case "SQRT": { double x = One(0); return x < 0 ? Value.Error("#NUM!") : Value.Of(Math.Sqrt(x)); }
                case "INT": return Value.Of(Math.Floor(One(0)));
                case "SIGN": return Value.Of(Math.Sign(One(0)));
                case "ROUND": return Value.Of(Math.Round(One(0), (int)One(1), MidpointRounding.AwayFromZero));
                case "ROUNDUP": { double f = Math.Pow(10, One(1)); double v = One(0); return Value.Of(Math.Ceiling(Math.Abs(v) * f) / f * Math.Sign(v)); }
                case "ROUNDDOWN": { double f = Math.Pow(10, One(1)); double v = One(0); return Value.Of(Math.Floor(Math.Abs(v) * f) / f * Math.Sign(v)); }
                case "MOD": { double b = One(1); return b == 0 ? Value.Error("#DIV/0!") : Value.Of(One(0) - b * Math.Floor(One(0) / b)); }
                case "POWER": return Value.Of(Math.Pow(One(0), One(1)));

                case "IF":
                {
                    var test = Flatten(args[0]);
                    if (test.IsError) return test;
                    if (!test.AsNumber(out var truth)) return Value.Error("#VALUE!");
                    if (truth != 0) return args.Count > 1 ? Flatten(args[1]) : Value.Of(true);
                    return args.Count > 2 ? Flatten(args[2]) : Value.Of(false);
                }
                case "AND": return Value.Of(All().All(v => v.AsNumber(out var n) && n != 0));
                case "OR": return Value.Of(All().Any(v => v.AsNumber(out var n) && n != 0));
                case "NOT": return Value.Of(One(0) == 0);
                case "IFERROR":
                {
                    var first = Flatten(args[0]);
                    return first.IsError ? Flatten(args[1]) : first;
                }

                case "LEN": return Value.Of(Str(0).Length);
                case "UPPER": return Value.Of(Str(0).ToUpperInvariant());
                case "LOWER": return Value.Of(Str(0).ToLowerInvariant());
                case "TRIM": return Value.Of(Str(0).Trim());
                case "LEFT": return Value.Of(Cut(Str(0), 0, args.Count > 1 ? (int)One(1) : 1));
                case "RIGHT": { var t = Str(0); int n = args.Count > 1 ? (int)One(1) : 1; return Value.Of(Cut(t, Math.Max(0, t.Length - n), n)); }
                case "MID": { var t = Str(0); int from = (int)One(1) - 1; return Value.Of(Cut(t, from, (int)One(2))); }
                case "CONCATENATE": return Value.Of(string.Concat(All().Select(v => v.AsText())));

                // ---- looking things up ----
                case "VLOOKUP" or "HLOOKUP":
                {
                    bool vertical = name == "VLOOKUP";
                    var needle = Flatten(args[0]);
                    var table = RangeOf(args[1]);
                    int index = (int)One(2) - 1;
                    bool approximate = args.Count < 4 || (Flatten(args[3]).AsNumber(out var flag) && flag != 0);
                    if (index < 0) return Value.Error("#VALUE!");

                    int length = vertical ? Rows(table) : Columns(table);
                    int best = -1;
                    for (int i = 0; i < length; i++)
                    {
                        var candidate = vertical ? At(table, i, 0) : At(table, 0, i);
                        if (!approximate)
                        {
                            if (Same(candidate, needle)) { best = i; break; }
                            continue;
                        }
                        // Approximate: the last one not greater than what is wanted, as the table is meant to be sorted.
                        if (Compare(candidate, needle, "<=").Number != 0) best = i; else break;
                    }
                    if (best < 0) return Value.Error("#N/A");
                    var found = vertical ? At(table, best, index) : At(table, index, best);
                    return found.Kind == Value.Sort.Blank ? Value.Of(0) : found;
                }

                case "XLOOKUP":
                {
                    var needle = Flatten(args[0]);
                    var haystack = RangeOf(args[1]);
                    var results = RangeOf(args[2]);
                    int length = Math.Max(Rows(haystack), Columns(haystack));
                    bool down = Rows(haystack) >= Columns(haystack);
                    for (int i = 0; i < length; i++)
                    {
                        var candidate = down ? At(haystack, i, 0) : At(haystack, 0, i);
                        if (!Same(candidate, needle)) continue;
                        return Rows(results) >= Columns(results) ? At(results, i, 0) : At(results, 0, i);
                    }
                    return args.Count > 3 ? Flatten(args[3]) : Value.Error("#N/A");
                }

                case "MATCH":
                {
                    var needle = Flatten(args[0]);
                    var haystack = RangeOf(args[1]);
                    int kind = args.Count > 2 ? (int)One(2) : 1;
                    int length = Math.Max(Rows(haystack), Columns(haystack));
                    bool down = Rows(haystack) >= Columns(haystack);
                    int best = -1;
                    for (int i = 0; i < length; i++)
                    {
                        var candidate = down ? At(haystack, i, 0) : At(haystack, 0, i);
                        if (kind == 0) { if (Same(candidate, needle)) return Value.Of(i + 1); continue; }
                        if (kind > 0) { if (Compare(candidate, needle, "<=").Number != 0) best = i; else break; }
                        else { if (Compare(candidate, needle, ">=").Number != 0) best = i; else break; }
                    }
                    return best < 0 ? Value.Error("#N/A") : Value.Of(best + 1);
                }

                case "INDEX":
                {
                    var table = RangeOf(args[0]);
                    int row = (int)One(1);
                    int column = args.Count > 2 ? (int)One(2) : 1;
                    // A single row or column can be indexed with one number.
                    if (args.Count == 2 && Rows(table) == 1) { column = row; row = 1; }
                    if (row < 1 || column < 1) return Value.Error("#VALUE!");
                    return At(table, row - 1, column - 1);
                }

                // ---- totals with a condition ----
                case "SUMIF" or "AVERAGEIF" or "COUNTIF":
                {
                    var over = RangeOf(args[0]);
                    var criterion = Flatten(args[1]);
                    var adding = name == "COUNTIF" ? over : args.Count > 2 ? RangeOf(args[2]) : over;
                    double total = 0; int matched = 0;
                    for (int r = 0; r < Rows(over); r++)
                        for (int c = 0; c < Columns(over); c++)
                        {
                            if (!Meets(At(over, r, c), criterion)) continue;
                            matched++;
                            if (name == "COUNTIF") continue;
                            if (At(adding, r, c).AsNumber(out var n)) total += n;
                        }
                    if (name == "COUNTIF") return Value.Of(matched);
                    if (name == "SUMIF") return Value.Of(total);
                    return matched == 0 ? Value.Error("#DIV/0!") : Value.Of(total / matched);
                }

                case "SUMIFS" or "COUNTIFS" or "AVERAGEIFS":
                {
                    bool counting = name == "COUNTIFS";
                    var adding = counting ? RangeOf(args[0]) : RangeOf(args[0]);
                    int first = counting ? 0 : 1;
                    var pairs = new List<(RefRange Over, Value Criterion)>();
                    for (int i = first; i + 1 < args.Count; i += 2)
                        pairs.Add((RangeOf(args[i]), Flatten(args[i + 1])));
                    if (pairs.Count == 0) throw new NotSupportedException();

                    var shape = pairs[0].Over;
                    double total = 0; int matched = 0;
                    for (int r = 0; r < Rows(shape); r++)
                        for (int c = 0; c < Columns(shape); c++)
                        {
                            if (!pairs.All(p => Meets(At(p.Over, r, c), p.Criterion))) continue;
                            matched++;
                            if (counting) continue;
                            if (At(adding, r, c).AsNumber(out var n)) total += n;
                        }
                    if (counting) return Value.Of(matched);
                    if (name == "SUMIFS") return Value.Of(total);
                    return matched == 0 ? Value.Error("#DIV/0!") : Value.Of(total / matched);
                }

                // ---- dates ----
                case "TODAY": return Value.Of(Serial(DateTime.Today));
                case "NOW": return Value.Of((DateTime.Now - Epoch).TotalDays);
                case "DATE": return Value.Of(Serial(new DateTime((int)One(0), 1, 1).AddMonths((int)One(1) - 1).AddDays(One(2) - 1)));
                case "YEAR": return Value.Of(FromSerial(One(0)).Year);
                case "MONTH": return Value.Of(FromSerial(One(0)).Month);
                case "DAY": return Value.Of(FromSerial(One(0)).Day);
                case "WEEKDAY": return Value.Of((int)FromSerial(One(0)).DayOfWeek + 1);
                case "DAYS": return Value.Of(One(0) - One(1));
                case "EDATE": return Value.Of(Serial(FromSerial(One(0)).AddMonths((int)One(1))));
                case "EOMONTH":
                {
                    var when = FromSerial(One(0)).AddMonths((int)One(1));
                    return Value.Of(Serial(new DateTime(when.Year, when.Month, DateTime.DaysInMonth(when.Year, when.Month))));
                }

                default: throw new NotSupportedException();   // a function Plain has never met
            }
        }
        catch (ErrorValue carried) { return carried.Value; }
        catch (ArgumentOutOfRangeException) { throw new NotSupportedException(); }
        catch (IndexOutOfRangeException) { throw new NotSupportedException(); }
    }

    private static string Cut(string text, int from, int length)
    {
        if (length < 0 || from < 0) return "";
        from = Math.Min(from, text.Length);
        return text.Substring(from, Math.Min(length, text.Length - from));
    }

    private sealed class ErrorValue : Exception
    {
        public Value Value { get; }
        public ErrorValue(Value value) { Value = value; }
    }

    /// <summary>Two values equal for the purpose of a lookup: numbers by value, text without regard to case.</summary>
    private static bool Same(Value a, Value b)
    {
        if (a.Kind is Value.Sort.Text || b.Kind is Value.Sort.Text)
            return string.Equals(a.AsText(), b.AsText(), StringComparison.CurrentCultureIgnoreCase);
        return a.AsNumber(out var x) && b.AsNumber(out var y) && x == y;
    }

    private static Value Arithmetic(Value a, Value b, char op)
    {
        if (a.IsError) return a;
        if (b.IsError) return b;
        if (!a.AsNumber(out var x) || !b.AsNumber(out var y)) return Value.Error("#VALUE!");
        return op switch
        {
            '+' => Value.Of(x + y),
            '-' => Value.Of(x - y),
            '*' => Value.Of(x * y),
            '/' => y == 0 ? Value.Error("#DIV/0!") : Value.Of(x / y),
            _ => Value.Error("#VALUE!"),
        };
    }

    private static Value Compare(Value a, Value b, string op)
    {
        if (a.IsError) return a;
        if (b.IsError) return b;

        int order;
        if (a.Kind is Value.Sort.Text || b.Kind is Value.Sort.Text)
            order = string.Compare(a.AsText(), b.AsText(), StringComparison.CurrentCultureIgnoreCase);
        else
        {
            a.AsNumber(out var x);
            b.AsNumber(out var y);
            order = x.CompareTo(y);
        }

        return Value.Of(op switch
        {
            "=" => order == 0,
            "<>" => order != 0,
            "<" => order < 0,
            ">" => order > 0,
            "<=" => order <= 0,
            _ => order >= 0,
        });
    }
}
