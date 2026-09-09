namespace Plain.Core;

/// <summary>Which way a sheet is being changed, and where.</summary>
public enum GridEdit { InsertRow, DeleteRow, InsertColumn, DeleteColumn }

/// <summary>
/// Moving the shape of a sheet: adding a line to a quote, taking a duplicate row out of a submission. This is the
/// thing people say they cannot finish a day's work without, and the hard part is not the cells, it is the formulas:
/// a total that read B2:B4 must read B2:B5 after a row goes in, or the number it shows becomes a lie.
///
/// A reference that pointed at something deleted becomes #REF!, which is what Excel does and what a person needs to
/// see, rather than a silently wrong answer.
/// </summary>
public static class Grid
{
    /// <summary>Where a cell ends up after the change, or null when the change removed it.</summary>
    public static CellRef? Move(CellRef cell, GridEdit edit, int at) => edit switch
    {
        GridEdit.InsertRow => cell.Row >= at ? new CellRef(cell.Column, cell.Row + 1) : cell,
        GridEdit.DeleteRow => cell.Row == at ? null : cell.Row > at ? new CellRef(cell.Column, cell.Row - 1) : cell,
        GridEdit.InsertColumn => cell.Column >= at ? new CellRef(cell.Column + 1, cell.Row) : cell,
        GridEdit.DeleteColumn => cell.Column == at ? null : cell.Column > at ? new CellRef(cell.Column - 1, cell.Row) : cell,
        _ => cell,
    };

    /// <summary>
    /// Rewrite a formula so it still means what it meant. A range that spanned the place a row went in grows to keep
    /// spanning it; a range whose every row was deleted has nothing left to point at and becomes #REF!.
    /// </summary>
    /// <param name="ownSheet">True when the formula lives on the sheet being changed, so its plain references are its own.</param>
    public static string Adjust(string formula, GridEdit edit, int at, string? onSheet, bool ownSheet = true)
    {
        return Refs.Rewrite(formula, token =>
        {
            var range = token.Range;
            // A plain reference belongs to the sheet the formula is on; a qualified one only moves when it names the
            // sheet being changed.
            bool mine = range.Sheet is null
                ? ownSheet
                : string.Equals(range.Sheet, onSheet, StringComparison.CurrentCultureIgnoreCase);
            if (!mine) return null;

            bool rows = edit is GridEdit.InsertRow or GridEdit.DeleteRow;
            int min = rows ? range.RowMin : range.ColumnMin;
            int max = rows ? range.RowMax : range.ColumnMax;
            int limit = rows ? 1048576 : 16384;

            int newMin = min, newMax = max;
            switch (edit)
            {
                case GridEdit.InsertRow or GridEdit.InsertColumn:
                    if (min >= at) newMin = min + 1;
                    // A range that ends at or after the insertion point stretches to keep covering what it covered.
                    if (max >= at) newMax = max + 1;
                    if (newMax > limit) newMax = limit;
                    if (newMin > limit) return "#REF!";
                    break;

                case GridEdit.DeleteRow or GridEdit.DeleteColumn:
                    if (min == max && min == at) return "#REF!";        // it pointed only at what went
                    if (min > at) newMin = min - 1;
                    if (max >= at) newMax = max - 1;
                    if (newMax < newMin) return "#REF!";
                    break;
            }

            if (newMin == min && newMax == max) return null;

            var moved = rows
                ? range with { RowMin = newMin, RowMax = newMax }
                : range with { ColumnMin = newMin, ColumnMax = newMax };
            return Refs.Write(moved, token.Kind, token.Column1Fixed, token.Row1Fixed,
                              token.Column2Fixed, token.Row2Fixed, token.IsRange);
        });
    }
}
