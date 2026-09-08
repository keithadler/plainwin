namespace Plain.Core;

public enum FileKind { Unknown, Spreadsheet, Document, Presentation }

/// <summary>
/// One Word, Excel or PowerPoint file, open. This is what the app and the console twin both talk to: it works out
/// which of the three it is, hands back the right model, and can always say what it is preserving.
/// </summary>
public sealed class PlainFile
{
    public OpcPackage Package { get; }
    public FileKind Kind { get; }
    public string Path { get; }

    public Workbook? Workbook { get; }
    public Document? Document { get; }
    public Deck? Deck { get; }

    private PlainFile(string path, OpcPackage pkg)
    {
        Path = path;
        Package = pkg;
        if (Core.Workbook.Looks(pkg)) { Kind = FileKind.Spreadsheet; Workbook = new Workbook(pkg); }
        else if (Core.Document.Looks(pkg)) { Kind = FileKind.Document; Document = new Document(pkg); }
        else if (Core.Deck.Looks(pkg)) { Kind = FileKind.Presentation; Deck = new Deck(pkg); }
        else throw new OpcPackage.PackageException("Plain opens Word, Excel and PowerPoint files. This is a package, but not one of those.");
    }

    public static PlainFile Open(string path) => new(path, OpcPackage.Open(path));

    /// <summary>The parts Plain draws on screen, as opposed to the ones it only keeps.</summary>
    public IEnumerable<string> ShownParts()
    {
        switch (Kind)
        {
            case FileKind.Spreadsheet:
                yield return Workbook!.WorkbookPart;
                if (Package.Has("xl/sharedStrings.xml")) yield return "xl/sharedStrings.xml";
                if (Package.Has("xl/styles.xml")) yield return "xl/styles.xml";
                foreach (var s in Workbook.Sheets) yield return s.PartName;
                break;
            case FileKind.Document:
                yield return Core.Document.BodyPart;
                break;
            case FileKind.Presentation:
                yield return "ppt/presentation.xml";
                foreach (var s in Deck!.Slides) yield return s.PartName;
                break;
        }
    }

    /// <summary>
    /// Every part with its role. Pending edits are pushed into the package first, because a part is only "edited"
    /// once the model has written it there, and a count taken before that would report a save as changing nothing.
    /// </summary>
    public IReadOnlyList<PartNote> Parts()
    {
        Flush();
        return Preserved.Describe(Package, ShownParts());
    }

    /// <summary>How many parts Plain is holding untouched, which is the number the status bar shows.</summary>
    public (int Total, int Edited, int Kept) Counts()
    {
        var parts = Parts();
        int edited = parts.Count(p => p.Role == PartRole.Edited);
        return (parts.Count, edited, parts.Count - edited);
    }

    public void Flush()
    {
        Workbook?.Flush();
        Document?.Flush();
        Deck?.Flush();
    }

    public void Save(string? path = null)
    {
        Flush();
        Package.Save(path ?? Path);
    }

    public static FileKind KindOf(string path) => System.IO.Path.GetExtension(path).ToLowerInvariant() switch
    {
        ".xlsx" or ".xlsm" => FileKind.Spreadsheet,
        ".docx" or ".docm" => FileKind.Document,
        ".pptx" or ".pptm" => FileKind.Presentation,
        _ => FileKind.Unknown,
    };
}
