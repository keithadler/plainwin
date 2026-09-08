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

    /// <summary>What the file looked like on disk when Plain opened it, so it can tell if something else changed it.</summary>
    public DateTime OpenedWritten { get; private set; }
    public long OpenedLength { get; private set; }

    private PlainFile(string path, OpcPackage pkg)
    {
        Path = path;
        Package = pkg;
        Remember();
        if (Core.Workbook.Looks(pkg)) { Kind = FileKind.Spreadsheet; Workbook = new Workbook(pkg); }
        else if (Core.Document.Looks(pkg)) { Kind = FileKind.Document; Document = new Document(pkg); }
        else if (Core.Deck.Looks(pkg)) { Kind = FileKind.Presentation; Deck = new Deck(pkg); }
        else throw new OpcPackage.PackageException("Plain opens Word, Excel and PowerPoint files. This is a package, but not one of those.");
    }

    public static PlainFile Open(string path) => new(path, OpcPackage.Open(path));

    /// <summary>Open from bytes already in hand, for tests and for anything that never touches a disk.</summary>
    public static PlainFile Read(byte[] bytes) => new("", OpcPackage.Read(bytes));

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

    private void Remember()
    {
        try
        {
            if (Path.Length == 0 || !File.Exists(Path)) return;
            var info = new FileInfo(Path);
            OpenedWritten = info.LastWriteTimeUtc;
            OpenedLength = info.Length;
        }
        catch { }
    }

    /// <summary>
    /// True when something else has written to the file since Plain opened it. Saving would then throw away whatever
    /// that was, so the app asks first rather than quietly winning.
    /// </summary>
    public bool ChangedOnDisk()
    {
        try
        {
            if (Path.Length == 0 || !File.Exists(Path)) return false;
            var info = new FileInfo(Path);
            return info.LastWriteTimeUtc != OpenedWritten || info.Length != OpenedLength;
        }
        catch { return false; }
    }

    /// <summary>True when Windows has the file marked read only, which Save would fail on.</summary>
    public bool IsReadOnly()
    {
        try { return Path.Length > 0 && File.Exists(Path) && File.GetAttributes(Path).HasFlag(FileAttributes.ReadOnly); }
        catch { return false; }
    }

    public void Save(string? path = null)
    {
        Flush();
        Package.Save(path ?? Path);
        if (path is null || path == Path) Remember();
    }

    public static FileKind KindOf(string path) => System.IO.Path.GetExtension(path).ToLowerInvariant() switch
    {
        ".xlsx" or ".xlsm" => FileKind.Spreadsheet,
        ".docx" or ".docm" => FileKind.Document,
        ".pptx" or ".pptm" => FileKind.Presentation,
        _ => FileKind.Unknown,
    };
}
