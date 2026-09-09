using System.Xml.Linq;

namespace Plain.Core;

/// <summary>
/// The choices a cell will accept.
///
/// A sheet someone else built often has cells that only take certain values: a status column that must be one of
/// four words, a date that must be this year. Excel calls it data validation and shows it as a little arrow. Plain
/// does not draw the arrow, but it can read the rule, and a cell that silently refuses what you type without
/// saying why is worse than one that tells you what it wants.
///
/// Plain reads these; it does not write them. Making a rule is not the fifth of the job this app is for, and a
/// half-made rule is worse than none.
/// </summary>
public static class Choices
{
    /// <summary>The rule on a cell: what kind, what it allows, and the words the file offers as help.</summary>
    public sealed record Rule(string Where, string Kind, IReadOnlyList<string> Allowed, string Formula, string Says);

    /// <summary>Every rule on the sheet.</summary>
    public static IReadOnlyList<Rule> On(OpcPackage pkg, Sheet sheet)
    {
        var found = new List<Rule>();
        if (!pkg.Has(sheet.PartName)) return found;

        var doc = Xml.Parse(pkg.Read(sheet.PartName));
        var d = Dialect.Of(doc.Root!);

        foreach (var rule in doc.Root!.Descendants(d.Sheet + "dataValidation"))
        {
            var where = (string?)rule.Attribute("sqref") ?? "";
            var kind = (string?)rule.Attribute("type") ?? "any";
            var first = rule.Element(d.Sheet + "formula1")?.Value ?? "";

            var allowed = new List<string>();
            // A list written into the rule itself is quoted and comma separated; one written as a range is a
            // reference, and Plain says so rather than pretending to have read it.
            if (kind == "list" && first.StartsWith('"') && first.EndsWith('"') && first.Length >= 2)
                allowed.AddRange(first[1..^1].Split(',', StringSplitOptions.TrimEntries | StringSplitOptions.RemoveEmptyEntries));

            var says = (string?)rule.Attribute("promptTitle") is { Length: > 0 } title
                ? title + ". " + ((string?)rule.Attribute("prompt") ?? "")
                : (string?)rule.Attribute("prompt") ?? "";

            found.Add(new Rule(where, Kind(kind), allowed, first, says.Trim()));
        }
        return found;
    }

    /// <summary>The rule covering this cell, if any.</summary>
    public static Rule? At(OpcPackage pkg, Sheet sheet, CellRef cell)
    {
        foreach (var rule in On(pkg, sheet))
            foreach (var part in rule.Where.Split(' ', StringSplitOptions.RemoveEmptyEntries))
            {
                var ends = part.Replace("$", "").Split(':');
                if (!CellRef.TryParse(ends[0], out var from)) continue;
                var to = ends.Length > 1 && CellRef.TryParse(ends[1], out var second) ? second : from;

                if (cell.Column >= Math.Min(from.Column, to.Column) && cell.Column <= Math.Max(from.Column, to.Column)
                 && cell.Row >= Math.Min(from.Row, to.Row) && cell.Row <= Math.Max(from.Row, to.Row))
                    return rule;
            }
        return null;
    }

    private static string Kind(string type) => type switch
    {
        "list" => "one of a list",
        "whole" => "a whole number",
        "decimal" => "a number",
        "date" => "a date",
        "time" => "a time",
        "textLength" => "text of a certain length",
        "custom" => "whatever a formula says",
        _ => type,
    };
}
