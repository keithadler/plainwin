using System.IO;
using System.Text;
using System.Text.Json;
using Plain.Core;
// WPF brings in System.Windows.Shapes, where Path is a shape. In this file Path always means the file system one.
using Path = System.IO.Path;

namespace Plain;

/// <summary>
/// plain.exe: the same program as a console app. Exit codes are 0 fine, 1 something to look at, 2 a problem,
/// 64 a usage mistake, so a script can act on the answer without parsing English.
/// </summary>
public static class Cli
{
    public const string Version = "1.0.1";

    private static readonly string[] Verbs =
        { "info", "parts", "text", "cells", "get", "set", "new", "replace", "row", "column", "props", "pdf", "csv", "import", "count", "images", "apply", "changes", "comments", "roundtrip", "selftest", "version", "help", "--help", "-h", "--version" };

    public static bool IsVerb(string arg) => Verbs.Contains(arg, StringComparer.OrdinalIgnoreCase);

    /// <summary>Was this switch given? Switches are compared exactly, so --case is not --Case.</summary>
    private static bool Flag(IReadOnlyList<string> args, string name) => args.Contains(name);

    private const string Usage = """
        Plain for Windows - opens Word, Excel and PowerPoint files, edits the basics,
        and never damages what it doesn't understand.

          plain new <file.xlsx|.docx|.pptx>  make a new empty file of that kind
          plain info <file>                what the file is, and what Plain keeps untouched
          plain parts <file> [--json]      every part, and whether Plain shows it or preserves it
          plain text <file>                the text, as plain text
          plain cells <file> [sheet]       every filled cell, as reference<tab>value
          plain get <file> <ref>           one cell, or one block by number
          plain set <file> <ref> <value>   change one cell or block, then save in place
          plain replace <file> <find> <with> [--case] [--whole] [--dry-run]
                                            change every occurrence, everywhere in the file
          plain row <file> insert|delete <n> [sheet]     put a row in or take one out
          plain column <file> insert|delete <ref> [sheet] the same for a column, named A, B, C
          plain props <file> [--set Name=value]... [--strip]
                                            show, change or clear what the file says about itself
          plain pdf <file> [out.pdf]       write it out as a PDF
          plain csv <file> [sheet] [--formatted]   write a sheet out as comma separated values;
                                            numbers come out as numbers unless you ask for the screen
          plain import <file> <csv> [at]   read a csv into a sheet, starting at a cell (default A1)
          plain count <file>               words, characters and paragraphs
          plain images <file> [--save dir] the pictures inside it, and optionally write them out
          plain changes <file> [--accept N|--reject N|--accept-all|--reject-all]
          plain comments <file> [--remove N|--remove-all]
          plain apply <file> --script <s>  run many changes in one go; - reads them from the input
          plain roundtrip <file>...        prove a save changes nothing: byte compares the result
          plain selftest [suite]           run the built-in checks
          plain version

        A cell is A1. A block is a number from "plain text --numbered". A value that starts
        with = is a formula. Nothing is uploaded and nothing is sent anywhere, ever.
        Exit codes: 0 fine, 1 something to look at, 2 problem, 64 usage.

        Built by Keith Adler. Free, MIT. github.com/keithadler/plainwin
        """;

    public static int Run(string[] args, TextWriter o, TextWriter err)
    {
        try { Console.OutputEncoding = Encoding.UTF8; } catch { }

        var positional = args.Where(a => !a.StartsWith("--")).ToList();
        bool json = args.Contains("--json");
        bool numbered = args.Contains("--numbered");
        string verb = positional.FirstOrDefault()?.ToLowerInvariant() ?? "help";
        var rest = positional.Skip(1).ToList();

        try
        {
            switch (verb)
            {
                case "help" or "--help" or "-h":
                    o.WriteLine(Usage);
                    return 0;

                case "version" or "--version":
                    o.WriteLine($"Plain for Windows {Version}");
                    o.WriteLine("Built by Keith Adler. Free and MIT licensed.");
                    o.WriteLine("github.com/keithadler/plainwin  -  keithadler.github.io");
                    return 0;

                case "selftest":
                    return Core.Tests.SelfTest.Run(o, rest.FirstOrDefault(), args.Contains("--list"));

                case "new":
                {
                    if (rest.Count < 1) { err.WriteLine("new <file.xlsx|file.docx|file.pptx>"); return 64; }
                    var made = PlainFile.Create(rest[0]);
                    o.WriteLine($"made {Path.GetFileName(made.Path)}: an empty {Describe(made.Kind)}, {made.Parts().Count} parts");
                    return 0;
                }

                case "info":
                {
                    if (rest.Count < 1) { err.WriteLine("info <file>"); return 64; }
                    using var _ = default(IDisposable);
                    var file = PlainFile.Open(rest[0]);
                    var (total, edited, kept) = file.Counts();
                    if (json)
                    {
                        o.WriteLine(JsonSerializer.Serialize(new
                        {
                            path = file.Path,
                            kind = file.Kind.ToString().ToLowerInvariant(),
                            parts = total,
                            shown = file.Parts().Count(p => p.Role == PartRole.Shown),
                            preserved = kept,
                            sheets = file.Workbook?.Sheets.Select(s => s.Name).ToArray(),
                            blocks = file.Document?.BlockCount,
                            slides = file.Deck?.Slides.Count,
                        }, new JsonSerializerOptions { WriteIndented = true }));
                        return 0;
                    }
                    o.WriteLine($"{Path.GetFileName(file.Path)}  ({Describe(file.Kind)})");
                    switch (file.Kind)
                    {
                        case FileKind.Spreadsheet:
                            foreach (var s in file.Workbook!.Sheets)
                                o.WriteLine($"  sheet    {s.Name}{(s.Hidden ? " (hidden)" : "")}, used to {s.Extent}");
                            break;
                        case FileKind.Document:
                            o.WriteLine($"  blocks   {file.Document!.BlockCount}");
                            break;
                        case FileKind.Presentation:
                            foreach (var s in file.Deck!.Slides) o.WriteLine($"  slide {s.Number,-3} {s.Title()}");
                            break;
                    }
                    int shown = file.Parts().Count(p => p.Role == PartRole.Shown);
                    o.WriteLine($"  parts    {total} in the file: {shown} shown, {total - shown} preserved untouched");
                    return 0;
                }

                case "parts":
                {
                    if (rest.Count < 1) { err.WriteLine("parts <file>"); return 64; }
                    var file = PlainFile.Open(rest[0]);
                    var parts = file.Parts();
                    if (json)
                    {
                        o.WriteLine(JsonSerializer.Serialize(parts.Select(p => new
                        { name = p.Name, role = p.Role.ToString().ToLowerInvariant(), what = p.What, bytes = p.Bytes }),
                        new JsonSerializerOptions { WriteIndented = true }));
                        return 0;
                    }
                    foreach (var p in parts)
                        o.WriteLine($"{(p.Role == PartRole.Shown ? "shown    " : "preserved")}  {p.Size,9}  {p.Name}   {p.What}");
                    return 0;
                }

                case "text":
                {
                    if (rest.Count < 1) { err.WriteLine("text <file>"); return 64; }
                    var file = PlainFile.Open(rest[0]);
                    switch (file.Kind)
                    {
                        case FileKind.Document:
                            foreach (var b in file.Document!.Blocks())
                                o.WriteLine(numbered ? $"{b.Index}\t{b.Text}" : b.Text);
                            return 0;
                        case FileKind.Presentation:
                            foreach (var s in file.Deck!.Slides)
                            {
                                o.WriteLine($"--- slide {s.Number} ---");
                                foreach (var t in s.Texts()) foreach (var line in t.Lines) o.WriteLine(line);
                            }
                            return 0;
                        default:
                            foreach (var sheet in file.Workbook!.Sheets)
                            {
                                o.WriteLine($"--- {sheet.Name} ---");
                                foreach (var c in sheet.Cells()) o.WriteLine(c.Display);
                            }
                            return 0;
                    }
                }

                case "cells":
                {
                    if (rest.Count < 1) { err.WriteLine("cells <file> [sheet]"); return 64; }
                    var file = PlainFile.Open(rest[0]);
                    if (file.Workbook is null) { err.WriteLine("cells only works on a spreadsheet."); return 2; }
                    var sheets = rest.Count > 1
                        ? file.Workbook.Sheets.Where(s => s.Name.Equals(rest[1], StringComparison.OrdinalIgnoreCase)).ToList()
                        : file.Workbook.Sheets.ToList();
                    if (sheets.Count == 0) { err.WriteLine($"no sheet called \"{rest[1]}\"."); return 2; }
                    foreach (var sheet in sheets)
                        foreach (var c in sheet.Cells())
                            o.WriteLine($"{sheet.Name}!{c.Ref}\t{c.Formula ?? c.Display}");
                    return 0;
                }

                case "get":
                {
                    if (rest.Count < 2) { err.WriteLine("get <file> <ref>"); return 64; }
                    var file = PlainFile.Open(rest[0]);
                    if (file.Workbook is not null)
                    {
                        if (!CellRef.TryParse(rest[1], out var cell)) { err.WriteLine($"\"{rest[1]}\" is not a cell reference."); return 64; }
                        var c = file.Workbook.Sheets[0].Read(cell);
                        o.WriteLine(c.Formula ?? c.Display);
                        return c.IsEmpty && c.Formula is null ? 1 : 0;
                    }
                    if (file.Document is not null && int.TryParse(rest[1], out var index))
                    {
                        if (index < 0 || index >= file.Document.BlockCount) { err.WriteLine("no block with that number."); return 2; }
                        o.WriteLine(file.Document.Read(index).Text);
                        return 0;
                    }
                    err.WriteLine("get takes a cell like A1 on a spreadsheet, or a block number on a document.");
                    return 64;
                }

                case "set":
                {
                    if (rest.Count < 3) { err.WriteLine("set <file> <ref> <value>"); return 64; }
                    string value = string.Join(' ', rest.Skip(2));
                    var file = PlainFile.Open(rest[0]);
                    if (file.Workbook is not null)
                    {
                        if (!CellRef.TryParse(rest[1], out var cell)) { err.WriteLine($"\"{rest[1]}\" is not a cell reference."); return 64; }
                        file.Workbook.Sheets[0].Set(cell, value);
                    }
                    else if (file.Document is not null && int.TryParse(rest[1], out var index))
                    {
                        if (index < 0 || index >= file.Document.BlockCount) { err.WriteLine("no block with that number."); return 2; }
                        if (!file.Document.SetText(index, value))
                            err.WriteLine("note: that block mixed several formats, so it now carries the first one throughout.");
                    }
                    else { err.WriteLine("set takes a cell like A1 on a spreadsheet, or a block number on a document."); return 64; }

                    var (total, edited, kept) = file.Counts();
                    file.Save();
                    o.WriteLine($"saved {Path.GetFileName(file.Path)}: {edited} of {total} parts rewritten, {kept} kept byte for byte");
                    return 0;
                }

                case "replace":
                {
                    if (rest.Count < 3) { err.WriteLine("replace <file> <find> <with>"); return 64; }
                    var options = new Replace.Options(
                        MatchCase: Flag(args, "--case"),
                        WholeCell: Flag(args, "--whole"),
                        IncludeFormulas: !Flag(args, "--no-formulas"));
                    var file = PlainFile.Open(rest[0]);
                    var result = Replace.InFile(file, rest[1], rest[2], options);
                    if (!result.Any) { o.WriteLine($"\"{rest[1]}\" is not in {Path.GetFileName(file.Path)}; nothing changed."); return 1; }
                    if (Flag(args, "--dry-run"))
                    {
                        o.WriteLine($"would change {result.Occurrences} occurrence{(result.Occurrences == 1 ? "" : "s")} in {result.Cells} place{(result.Cells == 1 ? "" : "s")}; nothing written");
                        return 0;
                    }
                    var (total, edited, kept) = file.Counts();
                    file.Save();
                    o.WriteLine($"changed {result.Occurrences} occurrence{(result.Occurrences == 1 ? "" : "s")} in {result.Cells} place{(result.Cells == 1 ? "" : "s")}; {edited} of {total} parts rewritten, {kept} kept byte for byte");
                    return 0;
                }

                case "row" or "column":
                {
                    if (rest.Count < 3) { err.WriteLine($"{verb} <file> insert|delete <where> [sheet]"); return 64; }
                    string how = rest[1].ToLowerInvariant();
                    if (how is not ("insert" or "delete")) { err.WriteLine("insert or delete"); return 64; }

                    var file = PlainFile.Open(rest[0]);
                    if (file.Workbook is null) { err.WriteLine($"{verb} only works on a spreadsheet."); return 2; }
                    var sheet = rest.Count > 3
                        ? file.Workbook.Sheets.FirstOrDefault(x => x.Name.Equals(rest[3], StringComparison.OrdinalIgnoreCase))
                        : file.Workbook.Sheets[0];
                    if (sheet is null) { err.WriteLine($"no sheet called \"{rest[3]}\"."); return 2; }

                    int at;
                    if (verb == "row")
                    {
                        if (!int.TryParse(rest[2], out at) || at < 1) { err.WriteLine("a row is a number from 1."); return 64; }
                    }
                    else
                    {
                        // A column is named the way it is on screen: A, B, AA.
                        if (!CellRef.TryParse(rest[2] + "1", out var asCell)) { err.WriteLine("a column is a letter like A or AB."); return 64; }
                        at = asCell.Column;
                    }

                    var edit = (verb, how) switch
                    {
                        ("row", "insert") => GridEdit.InsertRow,
                        ("row", _) => GridEdit.DeleteRow,
                        (_, "insert") => GridEdit.InsertColumn,
                        _ => GridEdit.DeleteColumn,
                    };
                    int touched = file.Workbook.Apply(sheet, edit, at);
                    var (total, edited, kept) = file.Counts();
                    file.Save();
                    o.WriteLine($"{(how == "insert" ? "inserted" : "deleted")} {verb} {rest[2]} on {sheet.Name}: {touched} formula{(touched == 1 ? "" : "s")} adjusted, {edited} of {total} parts rewritten, {kept} kept byte for byte");
                    return 0;
                }

                case "props":
                {
                    if (rest.Count < 1) { err.WriteLine("props <file> [--set Name=value] [--strip]"); return 64; }
                    var file = PlainFile.Open(rest[0]);
                    var props = file.Properties;

                    var sets = args.Where(a => a.StartsWith("--set=", StringComparison.Ordinal)).Select(a => a[6..]).ToList();
                    for (int i = 0; i < args.Length - 1; i++) if (args[i] == "--set") sets.Add(args[i + 1]);

                    bool changed = false;
                    foreach (var pair in sets)
                    {
                        int eq = pair.IndexOf('=');
                        if (eq <= 0) { err.WriteLine($"--set wants Name=value, not \"{pair}\"."); return 64; }
                        var name = pair[..eq].Trim();
                        if (!props.Set(name, pair[(eq + 1)..])) { err.WriteLine($"there is no property called \"{name}\"."); return 64; }
                        changed = true;
                    }

                    if (Flag(args, "--strip"))
                    {
                        int cleared = props.Strip();
                        o.WriteLine($"cleared {cleared} propert{(cleared == 1 ? "y" : "ies")} that named someone");
                        changed = changed || cleared > 0;
                    }

                    if (changed)
                    {
                        var (total, edited, kept) = file.Counts();
                        file.Save();
                        o.WriteLine($"saved: {edited} of {total} parts rewritten, {kept} kept byte for byte");
                        return 0;
                    }

                    if (json)
                    {
                        o.WriteLine(JsonSerializer.Serialize(props.All().ToDictionary(p => p.Name, p => p.Value),
                            new JsonSerializerOptions { WriteIndented = true }));
                        return 0;
                    }
                    foreach (var (name, value) in props.All())
                        o.WriteLine($"{name,-22}{(value.Length == 0 ? "-" : value)}");
                    var revealing = props.Revealing();
                    if (revealing.Count > 0)
                        o.WriteLine($"\n{revealing.Count} of these name a person or an organisation: {string.Join(", ", revealing.Select(r => r.Name))}");
                    return 0;
                }

                case "pdf":
                {
                    if (rest.Count < 1) { err.WriteLine("pdf <file> [out.pdf]"); return 64; }
                    var file = PlainFile.Open(rest[0]);
                    var to = rest.Count > 1 ? rest[1] : Path.ChangeExtension(file.Path, ".pdf");
                    var result = PdfExport.Build(file, Path.GetFileNameWithoutExtension(file.Path));
                    File.WriteAllBytes(to, result.Bytes);
                    o.WriteLine($"wrote {Path.GetFileName(to)}: {result.Pages} page{(result.Pages == 1 ? "" : "s")}, {result.Bytes.Length / 1024} KB");
                    if (result.Warning is not null) err.WriteLine(result.Warning);
                    return result.Warning is null ? 0 : 1;
                }

                case "csv":
                {
                    if (rest.Count < 1) { err.WriteLine("csv <file> [sheet]"); return 64; }
                    var file = PlainFile.Open(rest[0]);
                    if (file.Workbook is null) { err.WriteLine("csv only works on a spreadsheet."); return 2; }
                    var sheet = rest.Count > 1
                        ? file.Workbook.Sheets.FirstOrDefault(x => x.Name.Equals(rest[1], StringComparison.OrdinalIgnoreCase))
                        : file.Workbook.Sheets[0];
                    if (sheet is null) { err.WriteLine($"no sheet called \"{rest[1]}\"."); return 2; }
                    o.Write(Csv.Write(sheet, ',', Flag(args, "--formatted")));
                    return 0;
                }

                case "import":
                {
                    if (rest.Count < 2) { err.WriteLine("import <file> <csv> [at]"); return 64; }
                    var file = PlainFile.Open(rest[0]);
                    if (file.Workbook is null) { err.WriteLine("import only works on a spreadsheet."); return 2; }
                    var at = rest.Count > 2 ? rest[2] : "A1";
                    if (!CellRef.TryParse(at, out var corner)) { err.WriteLine($"\"{at}\" is not a cell reference."); return 64; }
                    var text = rest[1] == "-" ? Console.In.ReadToEnd() : File.ReadAllText(rest[1]);
                    int written = Csv.Into(file.Workbook.Sheets[0], text, corner);
                    var (total, edited, kept) = file.Counts();
                    file.Save();
                    o.WriteLine($"read {written} value{(written == 1 ? "" : "s")} into {file.Workbook.Sheets[0].Name} from {corner}; {edited} of {total} parts rewritten, {kept} kept byte for byte");
                    return 0;
                }

                case "count":
                {
                    if (rest.Count < 1) { err.WriteLine("count <file>"); return 64; }
                    var tally = Counts.Of(PlainFile.Open(rest[0]));
                    if (json) { o.WriteLine(JsonSerializer.Serialize(tally, new JsonSerializerOptions { WriteIndented = true })); return 0; }
                    o.WriteLine($"{tally.Words,10:N0}  words");
                    o.WriteLine($"{tally.Characters,10:N0}  characters");
                    o.WriteLine($"{tally.CharactersWithoutSpaces,10:N0}  characters without spaces");
                    o.WriteLine($"{tally.Paragraphs,10:N0}  paragraphs");
                    return 0;
                }

                case "images":
                {
                    if (rest.Count < 1) { err.WriteLine("images <file> [--save <dir>]"); return 64; }
                    var file = PlainFile.Open(rest[0]);
                    var pictures = Media.In(file);
                    if (pictures.Count == 0) { o.WriteLine("no pictures in this file"); return 1; }

                    string? into = null;
                    for (int i = 0; i < args.Length - 1; i++) if (args[i] == "--save") into = args[i + 1];

                    foreach (var picture in pictures)
                    {
                        if (into is null) o.WriteLine($"{picture.Size,10}  {picture.Kind,-18}{picture.Part}");
                        else o.WriteLine($"wrote {Media.SaveTo(picture, into)}");
                    }
                    return 0;
                }

                case "changes":
                {
                    if (rest.Count < 1) { err.WriteLine("changes <file> [--accept N|--reject N|--accept-all|--reject-all]"); return 64; }
                    var file = PlainFile.Open(rest[0]);
                    int? one = null; bool accept = true, act = false;
                    for (int i = 0; i < args.Length; i++)
                    {
                        if (args[i] is "--accept" or "--reject")
                        {
                            accept = args[i] == "--accept"; act = true;
                            if (i + 1 < args.Length && int.TryParse(args[i + 1], out var n)) one = n;
                        }
                        else if (args[i] is "--accept-all" or "--reject-all") { accept = args[i] == "--accept-all"; act = true; }
                    }

                    if (!act)
                    {
                        var list = Annotations.Revisions(file);
                        if (list.Count == 0) { o.WriteLine("no tracked changes"); return 1; }
                        for (int i = 0; i < list.Count; i++)
                            o.WriteLine($"{i,3}  {(list[i].Inserted ? "added     " : "struck out")}  {list[i].Author,-20}{list[i].Text}");
                        return 0;
                    }

                    int settled = one is { } index
                        ? (Annotations.SettleRevision(file, index, accept) ? 1 : 0)
                        : accept ? Annotations.AcceptRevisions(file) : Annotations.RejectRevisions(file);
                    if (settled == 0) { err.WriteLine("nothing to settle"); return 1; }
                    file.Save();
                    o.WriteLine($"{(accept ? "accepted" : "turned down")} {settled} change{(settled == 1 ? "" : "s")}");
                    return 0;
                }

                case "comments":
                {
                    if (rest.Count < 1) { err.WriteLine("comments <file> [--remove N|--remove-all]"); return 64; }
                    var file = PlainFile.Open(rest[0]);
                    int? one = null; bool all = Flag(args, "--remove-all");
                    for (int i = 0; i < args.Length - 1; i++)
                        if (args[i] == "--remove" && int.TryParse(args[i + 1], out var n)) one = n;

                    if (one is null && !all)
                    {
                        var list = Annotations.Comments(file);
                        if (list.Count == 0) { o.WriteLine("no comments"); return 1; }
                        for (int i = 0; i < list.Count; i++)
                            o.WriteLine($"{i,3}  {list[i].Author,-20}{list[i].When,-14}{list[i].Text}");
                        return 0;
                    }

                    int removed = one is { } index ? (Annotations.RemoveComment(file, index) ? 1 : 0)
                                                   : Annotations.RemoveComments(file);
                    if (removed == 0) { err.WriteLine("nothing to remove"); return 1; }
                    file.Save();
                    o.WriteLine($"removed {removed} comment{(removed == 1 ? "" : "s")}");
                    return 0;
                }

                case "apply":
                {
                    if (rest.Count < 1) { err.WriteLine("apply <file> --script <script|-> "); return 64; }
                    string? script = null;
                    for (int i = 0; i < args.Length - 1; i++) if (args[i] == "--script") script = args[i + 1];
                    if (script is null) { err.WriteLine("apply needs --script <file>, or --script - to read them from the input"); return 64; }

                    var lines = (script == "-" ? Console.In.ReadToEnd() : File.ReadAllText(script))
                        .Split('\n').Select(l => l.TrimEnd('\r').Trim())
                        .Where(l => l.Length > 0 && !l.StartsWith('#')).ToList();

                    var file = PlainFile.Open(rest[0]);
                    int done = 0;
                    foreach (var line in lines)
                    {
                        var result = Apply(file, line, err);
                        if (result != 0) { err.WriteLine($"stopped at: {line}"); return result; }
                        done++;
                    }
                    if (Flag(args, "--dry-run")) { o.WriteLine($"{done} change{(done == 1 ? "" : "s")} understood; nothing written"); return 0; }

                    var (total, edited, kept) = file.Counts();
                    file.Save();
                    o.WriteLine($"{done} change{(done == 1 ? "" : "s")} applied in one pass; {edited} of {total} parts rewritten, {kept} kept byte for byte");
                    return 0;
                }

                case "roundtrip":
                {
                    if (rest.Count < 1) { err.WriteLine("roundtrip <file>..."); return 64; }
                    int bad = 0;
                    foreach (var path in rest)
                    {
                        try
                        {
                            var original = File.ReadAllBytes(path);
                            var pkg = OpcPackage.Read(original);
                            foreach (var p in pkg.Parts) _ = pkg.Read(p.Name);
                            var saved = pkg.ToBytes();
                            bool same = saved.Length == original.Length && saved.AsSpan().SequenceEqual(original);
                            o.WriteLine($"{(same ? "identical" : "DIFFERS  ")}  {pkg.Parts.Count,4} parts  {Path.GetFileName(path)}");
                            if (!same) bad++;
                        }
                        catch (Exception ex) { o.WriteLine($"FAILED     {Path.GetFileName(path)}: {ex.Message}"); bad++; }
                    }
                    return bad == 0 ? 0 : 1;
                }

                default:
                    err.WriteLine($"\"{verb}\" is not something plain does.");
                    err.WriteLine(Usage);
                    return 64;
            }
        }
        catch (OpcPackage.PackageException ex) { err.WriteLine(ex.Message); return 2; }
        catch (FormatException ex) { err.WriteLine(ex.Message); return 64; }
        catch (FileNotFoundException) { err.WriteLine("No file at that path."); return 2; }
        catch (DirectoryNotFoundException) { err.WriteLine("No folder at that path."); return 2; }
        catch (UnauthorizedAccessException) { err.WriteLine("Windows would not let Plain read or write that file."); return 2; }
        catch (IOException ex) { err.WriteLine($"The file could not be read or written: {ex.Message}"); return 2; }
    }

    /// <summary>
    /// One line of a script: the same verbs as the command line, without naming the file each time. Four hundred
    /// changes in one process rather than four hundred processes, which is what made this worth having.
    /// </summary>
    private static int Apply(PlainFile file, string line, TextWriter err)
    {
        var parts = Split(line);
        if (parts.Count == 0) return 0;
        string verb = parts[0].ToLowerInvariant();

        try
        {
            switch (verb)
            {
                case "set" when parts.Count >= 3:
                {
                    string value = string.Join(' ', parts.Skip(2));
                    if (file.Workbook is not null)
                    {
                        if (!CellRef.TryParse(parts[1], out var cell)) { err.WriteLine($"\"{parts[1]}\" is not a cell."); return 64; }
                        file.Workbook.Sheets[0].Set(cell, value);
                        return 0;
                    }
                    if (file.Document is not null && int.TryParse(parts[1], out var block))
                    {
                        file.Document.SetText(block, value);
                        return 0;
                    }
                    err.WriteLine("set wants a cell on a spreadsheet or a block number on a document.");
                    return 64;
                }

                case "replace" when parts.Count >= 3:
                    Replace.InFile(file, parts[1], parts[2], new Replace.Options());
                    return 0;

                case "row" or "column" when parts.Count >= 3:
                {
                    if (file.Workbook is null) { err.WriteLine($"{verb} only works on a spreadsheet."); return 2; }
                    string how = parts[1].ToLowerInvariant();
                    int at;
                    if (verb == "row") { if (!int.TryParse(parts[2], out at)) { err.WriteLine("a row is a number."); return 64; } }
                    else { if (!CellRef.TryParse(parts[2] + "1", out var c)) { err.WriteLine("a column is a letter."); return 64; } at = c.Column; }
                    var edit = (verb, how) switch
                    {
                        ("row", "insert") => GridEdit.InsertRow, ("row", _) => GridEdit.DeleteRow,
                        (_, "insert") => GridEdit.InsertColumn, _ => GridEdit.DeleteColumn,
                    };
                    file.Workbook.Apply(file.Workbook.Sheets[0], edit, at);
                    return 0;
                }

                case "props" when parts.Count >= 2:
                {
                    if (parts[1] == "--strip") { file.Properties.Strip(); return 0; }
                    int eq = parts[1].IndexOf('=');
                    if (eq <= 0) { err.WriteLine("props wants Name=value or --strip."); return 64; }
                    if (!file.Properties.Set(parts[1][..eq], parts[1][(eq + 1)..])) { err.WriteLine($"no property called \"{parts[1][..eq]}\"."); return 64; }
                    return 0;
                }

                default:
                    err.WriteLine($"\"{verb}\" is not something a script can do.");
                    return 64;
            }
        }
        catch (Exception ex) { err.WriteLine(ex.Message); return 2; }
    }

    /// <summary>Split a script line, letting quotes hold a value with spaces in it together.</summary>
    private static List<string> Split(string line)
    {
        var parts = new List<string>();
        var built = new System.Text.StringBuilder();
        bool quoted = false;
        foreach (char c in line)
        {
            if (c == '"') { quoted = !quoted; continue; }
            if (c == ' ' && !quoted) { if (built.Length > 0) { parts.Add(built.ToString()); built.Clear(); } continue; }
            built.Append(c);
        }
        if (built.Length > 0) parts.Add(built.ToString());
        return parts;
    }

    private static string Describe(FileKind kind) => kind switch
    {
        FileKind.Spreadsheet => "Excel workbook",
        FileKind.Document => "Word document",
        FileKind.Presentation => "PowerPoint deck",
        _ => "unknown",
    };
}
