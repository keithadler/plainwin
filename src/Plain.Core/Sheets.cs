using System.Xml.Linq;

namespace Plain.Core;

/// <summary>
/// Changing the shape of a workbook: adding a sheet, renaming one, taking one out, moving one.
///
/// A workbook keeps its sheets in one list in workbook.xml, each pointing at a worksheet part through a
/// relationship. Renaming and reordering only touch that list, so both are safe. Removing takes the entry out and
/// leaves the part in the package unused, for the same reason a removed slide keeps its part: deleting a part that
/// something unnoticed still points at is how a file gets broken.
///
/// Renaming is the one that needs care beyond the list. A formula that says Detail!B2 means the sheet called
/// Detail, so renaming a sheet without rewriting the formulas that name it turns them into references to a sheet
/// that no longer exists. Every formula in the workbook is rewritten to follow the new name.
/// </summary>
public static class Sheets
{
    private const string DocRel = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

    public abstract record Result;
    public sealed record Done(string What) : Result;
    public sealed record Refused(string Reason) : Result;

    /// <summary>Names Excel will not accept, and the length it will not go past.</summary>
    public static Result CheckName(Workbook book, string name, string? except = null)
    {
        var trimmed = name.Trim();
        if (trimmed.Length == 0) return new Refused("A sheet has to have a name.");
        if (trimmed.Length > 31) return new Refused("A sheet's name cannot be longer than 31 characters.");
        foreach (char c in ":\\/?*[]")
            if (trimmed.Contains(c)) return new Refused($"A sheet's name cannot contain {c}");
        if (trimmed.StartsWith('\'') || trimmed.EndsWith('\''))
            return new Refused("A sheet's name cannot start or end with an apostrophe.");
        if (book.Sheets.Any(s => !string.Equals(s.Name, except, StringComparison.OrdinalIgnoreCase)
                              && string.Equals(s.Name, trimmed, StringComparison.OrdinalIgnoreCase)))
            return new Refused($"There is already a sheet called {trimmed}.");
        return new Done(trimmed);
    }

    /// <summary>
    /// Rename a sheet, and rewrite every formula in the workbook that named it, so nothing points at a sheet that
    /// is no longer there.
    /// </summary>
    public static Result Rename(Workbook book, int index, string name)
    {
        if (index < 0 || index >= book.Sheets.Count) return new Refused($"There is no sheet {index + 1}.");
        var sheet = book.Sheets[index];
        if (CheckName(book, name, except: sheet.Name) is Refused bad) return bad;
        var wanted = name.Trim();
        if (wanted == sheet.Name) return new Done($"That sheet is already called {wanted}.");

        int touched = book.RenameSheet(index, wanted);
        return new Done(touched == 0
            ? $"Renamed to {wanted}."
            : $"Renamed to {wanted}, and {touched} formula{(touched == 1 ? "" : "s")} that named it now say so.");
    }

    /// <summary>Move a sheet to a new place in the running order, counting from 1.</summary>
    public static Result Move(Workbook book, int from, int to)
    {
        if (from < 1 || from > book.Sheets.Count) return new Refused($"There is no sheet {from}.");
        to = Math.Clamp(to, 1, book.Sheets.Count);
        if (from == to) return new Done("That sheet is already there.");
        book.MoveSheet(from - 1, to - 1);
        return new Done($"Sheet {from} is now sheet {to}.");
    }

    /// <summary>Take a sheet out. Its part stays in the file, unused, rather than being deleted.</summary>
    public static Result Remove(Workbook book, int index)
    {
        if (index < 0 || index >= book.Sheets.Count) return new Refused($"There is no sheet {index + 1}.");
        if (book.Sheets.Count <= 1) return new Refused("A workbook has to have a sheet in it, so this one cannot go.");

        var name = book.Sheets[index].Name;
        int pointing = book.FormulasNaming(name);
        if (pointing > 0)
            return new Refused($"{pointing} formula{(pointing == 1 ? "" : "s")} read {name}. "
                             + "Taking it out would leave them pointing at nothing, so Plain has left it alone.");

        book.RemoveSheet(index);
        return new Done($"{name} is out of the workbook. Its contents stay in the file, unused.");
    }

    /// <summary>Put a new empty sheet in after the one given.</summary>
    public static Result Add(Workbook book, string name, int after)
    {
        if (CheckName(book, name) is Refused bad) return bad;
        var wanted = name.Trim();
        book.AddSheet(wanted, after);
        return new Done($"{wanted} is in.");
    }
}
