using System.Xml.Linq;

namespace Plain.Core;

/// <summary>
/// What is in a file that you probably did not mean to send.
///
/// Office files carry a great deal you never see: who wrote them and who edited them last, comments arguing about
/// a number, tracked changes showing what a clause used to say, sheets and rows someone hid rather than deleted,
/// speaker notes meant for the person presenting. None of it is secret and all of it travels with the file.
///
/// This is the same idea as the preserved rail, pointed at the moment people get burned: it lists what is there,
/// in plain words, and takes out only what you choose. It never removes anything on its own, and it never claims
/// to have made a file safe, because it can only find the things it knows how to look for.
/// </summary>
public static class Hidden
{
    /// <summary>One kind of thing found, how many there are, and whether Plain can take it out.</summary>
    public sealed record Finding(string Kind, string What, int Count, bool CanRemove);

    public const string People = "Names and history";
    public const string Comments = "Comments";
    public const string Tracked = "Tracked changes";
    public const string HiddenSheets = "Hidden sheets";
    public const string HiddenRows = "Hidden rows";
    public const string Notes = "Speaker notes";

    /// <summary>Everything Plain can find. An empty list means nothing it knows to look for is there.</summary>
    public static IReadOnlyList<Finding> Find(PlainFile file)
    {
        var found = new List<Finding>();

        // ---- who wrote it, and when ----
        var naming = file.Properties.Revealing();
        if (naming.Count > 0)
            found.Add(new Finding(People,
                string.Join(", ", naming.Select(x => $"{x.Name}: {x.Value}")), naming.Count, true));

        // ---- comments and tracked changes ----
        int comments = Annotations.Comments(file).Count;
        if (comments > 0)
            found.Add(new Finding(Comments,
                $"{comments} comment{(comments == 1 ? "" : "s")}, with who wrote them and when", comments, true));

        int changes = Annotations.Revisions(file).Count;
        if (changes > 0)
            found.Add(new Finding(Tracked,
                $"{changes} tracked change{(changes == 1 ? "" : "s")}, showing what the text used to say", changes, true));

        // ---- sheets and rows someone hid ----
        if (file.Workbook is { } book)
        {
            var hiddenSheets = book.Sheets.Where(x => x.Hidden).Select(x => x.Name).ToList();
            if (hiddenSheets.Count > 0)
                found.Add(new Finding(HiddenSheets, "Hidden: " + string.Join(", ", hiddenSheets), hiddenSheets.Count, false));

            int rows = 0;
            foreach (var sheet in book.Sheets) rows += sheet.HiddenRowCount();
            if (rows > 0)
                found.Add(new Finding(HiddenRows, $"{rows} hidden row{(rows == 1 ? "" : "s")}", rows, false));
        }

        // ---- what the presenter was going to say ----
        if (file.Deck is not null)
        {
            int notes = file.Package.Parts.Count(p =>
                p.Name.StartsWith("ppt/notesSlides/notesSlide", StringComparison.Ordinal)
                && p.Name.EndsWith(".xml", StringComparison.Ordinal));
            if (notes > 0)
                found.Add(new Finding(Notes, $"{notes} slide{(notes == 1 ? "" : "s")} with speaker notes", notes, false));
        }

        return found;
    }

    /// <summary>
    /// Take out the kinds chosen. Only the kinds marked as removable can go: hidden sheets and rows are somebody's
    /// working, not an accident, and speaker notes live in parts of their own that Plain will not tear out of a
    /// deck. Says what it did, in the same words the finding used.
    /// </summary>
    public static IReadOnlyList<string> Remove(PlainFile file, IReadOnlyCollection<string> kinds)
    {
        var did = new List<string>();

        if (kinds.Contains(People))
        {
            int n = file.Properties.Strip();
            if (n > 0) did.Add($"Took out {n} propert{(n == 1 ? "y" : "ies")} that named somebody.");
        }

        if (kinds.Contains(Comments))
        {
            int n = Annotations.RemoveComments(file);
            if (n > 0) did.Add($"Took out {n} comment{(n == 1 ? "" : "s")}.");
        }

        if (kinds.Contains(Tracked))
        {
            int n = Annotations.AcceptRevisions(file);
            if (n > 0) did.Add($"Settled {n} tracked change{(n == 1 ? "" : "s")}, keeping what was added.");
        }

        return did;
    }
}
