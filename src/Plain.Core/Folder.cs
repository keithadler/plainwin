using Path = System.IO.Path;
// This class has a Search of its own, so the one that looks inside a file is named apart from it.
using Look = Plain.Core.Search;

namespace Plain.Core;

/// <summary>
/// Looking for something across every Office file in a folder.
///
/// The job this exists for is "which of these forty documents mentions the old company name". Doing that by hand
/// means opening forty files; doing it with a tool that also offers to change them means trusting it with forty
/// files at once. So this only ever reads. It says which files hold the words and where, and then you open the
/// ones that matter and change them yourself, one at a time, watching what happens.
/// </summary>
public static class Folder
{
    /// <summary>One file that holds the words, and the first few places it holds them.</summary>
    public sealed record Hit(string Path, int Count, IReadOnlyList<string> Places);

    /// <summary>A file that could not be read, and why, so a folder with one bad file still gives an answer.</summary>
    public sealed record Trouble(string Path, string Why);

    public sealed record Report(IReadOnlyList<Hit> Hits, IReadOnlyList<Trouble> Troubles, int Looked);

    public static readonly string[] Extensions = { ".docx", ".docm", ".xlsx", ".xlsm", ".pptx", ".pptm" };

    /// <summary>
    /// Search every Office file in the folder. <paramref name="deep"/> looks in folders inside it too.
    /// <paramref name="stopAfter"/> bounds the work so a search of a whole disk does not run for an hour.
    /// </summary>
    public static Report Search(string folder, string term, bool deep = false, int stopAfter = 2000,
                                int placesPerFile = 5, Func<bool>? cancelled = null)
    {
        var hits = new List<Hit>();
        var troubles = new List<Trouble>();
        int looked = 0;

        if (term.Trim().Length == 0 || !Directory.Exists(folder)) return new Report(hits, troubles, 0);

        IEnumerable<string> files;
        try
        {
            files = Directory.EnumerateFiles(folder, "*",
                deep ? SearchOption.AllDirectories : SearchOption.TopDirectoryOnly);
        }
        catch (Exception ex) { return new Report(hits, new[] { new Trouble(folder, Explain(ex)) }, 0); }

        foreach (var path in files)
        {
            if (cancelled?.Invoke() == true) break;
            if (looked >= stopAfter) break;

            var extension = Path.GetExtension(path);
            if (!Extensions.Contains(extension, StringComparer.OrdinalIgnoreCase)) continue;
            // Word leaves these behind while a file is open; they are not documents anyone means to search.
            if (Path.GetFileName(path).StartsWith("~$", StringComparison.Ordinal)) continue;

            looked++;
            try
            {
                var file = PlainFile.Open(path);
                var places = new List<string>();
                int count = 0;

                if (file.Workbook is { } book)
                    foreach (var hit in Look.InWorkbook(book, term))
                    { count++; if (places.Count < placesPerFile) places.Add($"{hit.Where}: {Trim(hit.Text)}"); }
                else if (file.Document is { } doc)
                    foreach (var hit in Look.InDocument(doc, term))
                    { count++; if (places.Count < placesPerFile) places.Add(Trim(hit.Text)); }
                else if (file.Deck is { } deck)
                    foreach (var hit in Look.InDeck(deck, term))
                    { count++; if (places.Count < placesPerFile) places.Add($"{hit.Where}: {Trim(hit.Text)}"); }

                if (count > 0) hits.Add(new Hit(path, count, places));
            }
            catch (Exception ex) { troubles.Add(new Trouble(path, Explain(ex))); }
        }

        return new Report(hits, troubles, looked);
    }

    private static string Trim(string text)
    {
        var one = text.Replace('\n', ' ').Replace('\r', ' ').Trim();
        return one.Length <= 90 ? one : one[..90] + "…";
    }

    private static string Explain(Exception ex) => ex switch
    {
        OpcPackage.PackageException p => p.Message,
        UnauthorizedAccessException => "Plain was not allowed to read it.",
        IOException => "Something else has it open.",
        _ => ex.Message,
    };
}
