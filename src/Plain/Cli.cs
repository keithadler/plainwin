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
    public const string Version = "1.4.2";

    private static readonly string[] Verbs =
        { "info", "parts", "text", "cells", "get", "set", "new", "replace", "row", "column", "width", "freeze", "sort", "explorer", "slide", "tablerow", "hidden", "find", "sheet", "height", "align", "colour", "band", "links", "compare", "border", "tidy", "reads", "choices", "notes", "describe", "link", "picture", "props", "pdf", "csv", "import", "count", "images", "apply", "changes", "comments", "roundtrip", "selftest", "version", "help", "--help", "-h", "--version" };

    public static bool IsVerb(string arg) => Verbs.Contains(arg, StringComparer.OrdinalIgnoreCase);

    /// <summary>Was this switch given? Switches are compared exactly, so --case is not --Case.</summary>
    private static bool Flag(IReadOnlyList<string> args, string name) => args.Contains(name);

    /// <summary>
    /// The switches that are followed by a value. Everything else is on or off. This list is what keeps a flag's
    /// value from being read as one of the words the verb was given.
    /// </summary>
    private static readonly HashSet<string> TakesAValue = new(StringComparer.Ordinal)
    {
        "--sheet", "--on", "--sides", "--colour", "--color", "--fill", "--ink",
        "--set", "--cm", "--remove", "--since",
    };

    /// <summary>The sheet by name, or the first one when no name was given.</summary>
    private static Sheet? PickSheet(Workbook book, string? name) =>
        name is null ? book.Sheets[0]
                     : book.Sheets.FirstOrDefault(x => x.Name.Equals(name, StringComparison.OrdinalIgnoreCase));

    /// <summary>Every cell in a range written as B2 or A2:D40.</summary>
    private static List<CellRef> Block(string range)
    {
        var cells = new List<CellRef>();
        var ends = range.Split(':');
        if (!CellRef.TryParse(ends[0], out var from)) return cells;
        var to = ends.Length > 1 && CellRef.TryParse(ends[1], out var second) ? second : from;
        for (int r = Math.Min(from.Row, to.Row); r <= Math.Max(from.Row, to.Row); r++)
            for (int c = Math.Min(from.Column, to.Column); c <= Math.Max(from.Column, to.Column); c++)
                cells.Add(new CellRef(c, r));
        return cells;
    }

    /// <summary>What followed this switch, or nothing if it was not given or was given nothing.</summary>
    private static string? Value(IReadOnlyList<string> args, string name)
    {
        for (int i = 0; i < args.Count - 1; i++)
            if (args[i] == name && !args[i + 1].StartsWith("--", StringComparison.Ordinal)) return args[i + 1];
        return null;
    }

    private const string Usage = """
        Plain for Windows - opens Word, Excel and PowerPoint files, edits the basics,
        and never damages what it doesn't understand.

          plain new <file.xlsx|.docx|.pptx>  make a new empty file of that kind
          plain width <file> <col> <chars|fit> [sheet]   set a column's width, or fit it to its contents
          plain freeze <file> <rows> [sheet]             keep this many rows at the top on screen
          plain sort <file> <range> <col> [down] [sheet] sort rows, refusing if a formula would be broken
          plain explorer on|off|status     "Edit in Plain" on the right-click menu, for your account only
          plain slide <file> add|remove|move <n> [to]   add, take out or move a slide
          plain tablerow <file> add|remove <table> <row>  a row in a table in a document
          plain hidden <file> [--remove k,k]  what travels with the file that you may not want to send
          plain find <folder> <words> [--deep]  which Office files in a folder hold those words
          plain sheet <file> add|rename|remove|move <n> [name|to]   sheets in a workbook
          plain height <file> <row> <points|auto> [sheet]           how tall a row is
          plain align <file> <range> left|centre|right|general [--wrap|--nowrap] [--sheet <name>]
          plain colour <file> <range> [--fill RRGGBB] [--ink RRGGBB] [--sheet <name>]
          plain band <file> [--set <n> "<text>"]   the lines along the top and bottom of every page
          plain links <file>               every link, and where it actually goes
          plain compare <before> <after>   what changed between two versions of a file
          plain border <file> <range> <weight|none> [--sides all|top,left] [--colour RRGGBB] [--sheet <n>]
          plain tidy <file> dedupe|split <range> [--on ","] [--sheet <n>]
          plain reads <file> <cell> [--sheet <n>]   what a formula reads, and what reads the cell
          plain choices <file> [--sheet <n>]        the choices cells will accept
          plain notes <file> [--set <slide> "<text>"]   what the presenter was going to say
          plain describe <file> [--set "<text>"]    the words that describe each picture
          plain link <file> <paragraph> <address>|off   put a link on a paragraph
          plain picture <file> <image> [--cm 8]     put a picture at the end of a document

        For pdf: --paper A4|Letter|Legal|A3|A5  --landscape  --margin <mm>
                 --header "<text>"  --footer "<text>"   with {page} and {pages}
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

        // A switch that takes a value swallows the word after it. Without this, "set book.xlsx B5 25000
        // --sheet Summary" put the text "25000 Summary" in the cell: a number quietly became words, and a
        // column that no longer adds up is exactly what this program exists to prevent.
        var positional = new List<string>();
        for (int i = 0; i < args.Length; i++)
        {
            if (args[i].StartsWith("--", StringComparison.Ordinal))
            {
                if (TakesAValue.Contains(args[i]) && i + 1 < args.Length
                    && !args[i + 1].StartsWith("--", StringComparison.Ordinal)) i++;
                continue;
            }
            positional.Add(args[i]);
        }
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

                case "slide":
                {
                    if (rest.Count < 2) { err.WriteLine("slide <file> add|remove|move <n> [to]"); return 64; }
                    var file = PlainFile.Open(rest[0]);
                    if (file.Deck is null) { err.WriteLine("slide only works on a presentation."); return 2; }

                    string how = rest[1].ToLowerInvariant();
                    int at = rest.Count > 2 && int.TryParse(rest[2], out var n) ? n : 0;
                    Slides.Result outcome = how switch
                    {
                        "add" => Slides.Add(file.Package, at),
                        "remove" => Slides.Remove(file.Package, at),
                        "move" => rest.Count > 3 && int.TryParse(rest[3], out var to)
                                  ? Slides.Move(file.Package, at, to)
                                  : new Slides.Refused("move needs the slide and where it goes, as in: move 3 1"),
                        _ => new Slides.Refused("add, remove or move"),
                    };

                    if (outcome is Slides.Refused refused) { err.WriteLine(refused.Reason); return 2; }
                    file.Flush();
                    File.WriteAllBytes(file.Path, file.Package.ToBytes());
                    o.WriteLine(((Slides.Done)outcome).What);
                    return 0;
                }

                case "tablerow":
                {
                    if (rest.Count < 4) { err.WriteLine("tablerow <file> add|remove <table> <row>"); return 64; }
                    var file = PlainFile.Open(rest[0]);
                    if (file.Document is null) { err.WriteLine("tablerow only works on a document."); return 2; }
                    if (!int.TryParse(rest[2], out var table) || !int.TryParse(rest[3], out var row))
                    { err.WriteLine("the table and the row are numbers, counting from 1."); return 64; }

                    var outcome = rest[1].ToLowerInvariant() switch
                    {
                        "add" => file.Document.InsertRow(table - 1, row),
                        "remove" => file.Document.DeleteRow(table - 1, row),
                        _ => new TableRows.Refused("add or remove"),
                    };
                    if (outcome is TableRows.Refused refused) { err.WriteLine(refused.Reason); return 2; }
                    file.Flush();
                    File.WriteAllBytes(file.Path, file.Package.ToBytes());
                    o.WriteLine(((TableRows.Done)outcome).What);
                    return 0;
                }

                case "hidden":
                {
                    if (rest.Count < 1) { err.WriteLine("hidden <file> [--remove comments,tracked,names]"); return 64; }
                    var file = PlainFile.Open(rest[0]);
                    var found = Hidden.Find(file);
                    if (found.Count == 0)
                    {
                        o.WriteLine("Nothing found that you would not expect to send.");
                        o.WriteLine("Plain can only find what it knows to look for, so this is not a promise.");
                        return 0;
                    }

                    foreach (var f in found)
                        o.WriteLine($"{f.Kind,-20} {f.What}" + (f.CanRemove ? "" : "   (Plain leaves this alone)"));

                    var wanted = Value(args, "--remove");
                    if (wanted is null) { o.WriteLine(); o.WriteLine("Use --remove names,comments,tracked to take them out."); return 1; }

                    var kinds = new List<string>();
                    foreach (var word in wanted.Split(',', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries))
                        kinds.Add(word.ToLowerInvariant() switch
                        {
                            "names" or "people" or "properties" => Hidden.People,
                            "comments" => Hidden.Comments,
                            "tracked" or "changes" => Hidden.Tracked,
                            _ => word,
                        });

                    var did = Hidden.Remove(file, kinds);
                    if (did.Count == 0) { err.WriteLine("Nothing was taken out. The kinds are: names, comments, tracked."); return 2; }
                    file.Flush();
                    File.WriteAllBytes(file.Path, file.Package.ToBytes());
                    o.WriteLine();
                    foreach (var line in did) o.WriteLine(line);
                    return 0;
                }

                case "find":
                {
                    if (rest.Count < 2) { err.WriteLine("find <folder> <words> [--deep]"); return 64; }
                    var report = Folder.Search(rest[0], string.Join(' ', rest.Skip(1)), Flag(args, "--deep"));

                    foreach (var hit in report.Hits)
                    {
                        o.WriteLine($"{hit.Path}  ({hit.Count})");
                        foreach (var place in hit.Places) o.WriteLine("    " + place);
                    }
                    foreach (var trouble in report.Troubles)
                        err.WriteLine($"{trouble.Path}: {trouble.Why}");

                    o.WriteLine();
                    o.WriteLine($"{report.Hits.Count} of {report.Looked} files hold it. Nothing was changed.");
                    return report.Hits.Count > 0 ? 0 : 1;
                }

                case "sheet":
                {
                    if (rest.Count < 2) { err.WriteLine("sheet <file> add|rename|remove|move <n> [name|to]"); return 64; }
                    var file = PlainFile.Open(rest[0]);
                    if (file.Workbook is null) { err.WriteLine("sheet only works on a spreadsheet."); return 2; }
                    var book = file.Workbook;

                    string how = rest[1].ToLowerInvariant();
                    int at = rest.Count > 2 && int.TryParse(rest[2], out var n) ? n : 0;
                    Sheets.Result outcome = how switch
                    {
                        "add" => Sheets.Add(book, rest.Count > 2 ? rest[2] : "Sheet", book.Sheets.Count),
                        "rename" => rest.Count > 3
                            ? Sheets.Rename(book, at - 1, rest[3])
                            : new Sheets.Refused("rename needs the sheet and the new name: rename 2 Invoices"),
                        "remove" => Sheets.Remove(book, at - 1),
                        "move" => rest.Count > 3 && int.TryParse(rest[3], out var to)
                            ? Sheets.Move(book, at, to)
                            : new Sheets.Refused("move needs the sheet and where it goes: move 3 1"),
                        _ => new Sheets.Refused("add, rename, remove or move"),
                    };
                    if (outcome is Sheets.Refused refused) { err.WriteLine(refused.Reason); return 2; }
                    file.Flush();
                    File.WriteAllBytes(file.Path, file.Package.ToBytes());
                    o.WriteLine(((Sheets.Done)outcome).What);
                    return 0;
                }

                case "height":
                {
                    if (rest.Count < 3) { err.WriteLine("height <file> <row> <points|auto> [sheet]"); return 64; }
                    var file = PlainFile.Open(rest[0]);
                    if (file.Workbook is null) { err.WriteLine("height only works on a spreadsheet."); return 2; }
                    var sheet = PickSheet(file.Workbook, rest.Count > 3 ? rest[3] : null);
                    if (sheet is null) { err.WriteLine($"no sheet called \"{rest[3]}\"."); return 2; }
                    if (!int.TryParse(rest[1], out var row) || row < 1) { err.WriteLine("a row is a number from 1."); return 64; }

                    double points = rest[2].Equals("auto", StringComparison.OrdinalIgnoreCase) ? 0
                        : double.TryParse(rest[2], System.Globalization.NumberStyles.Float,
                                          System.Globalization.CultureInfo.InvariantCulture, out var p) ? p : -1;
                    if (points < 0) { err.WriteLine("a height is a number of points, or the word auto."); return 64; }

                    sheet.SetHeightPoints(row, points);
                    file.Flush();
                    File.WriteAllBytes(file.Path, file.Package.ToBytes());
                    o.WriteLine(points <= 0
                        ? $"row {row} on {sheet.Name} follows the sheet again"
                        : $"row {row} on {sheet.Name} is now {sheet.HeightPoints(row)} points tall");
                    return 0;
                }

                case "align" or "colour":
                {
                    if (rest.Count < 2) { err.WriteLine($"{verb} <file> <range> ..."); return 64; }
                    var file = PlainFile.Open(rest[0]);
                    if (file.Workbook is null) { err.WriteLine($"{verb} only works on a spreadsheet."); return 2; }
                    // These verbs take switches, and a switch's value stays in the positional list, so the sheet
                    // is named with --sheet rather than by position; "colour f A1 --fill FFE7A1" would otherwise
                    // read the colour as a sheet name.
                    var wantedSheet = Value(args, "--sheet");
                    var sheet = PickSheet(file.Workbook, wantedSheet);
                    if (sheet is null) { err.WriteLine($"no sheet called \"{wantedSheet}\"."); return 2; }

                    var cells = Block(rest[1]);
                    if (cells.Count == 0) { err.WriteLine("a range looks like B2 or A2:D40."); return 64; }

                    if (verb == "align")
                    {
                        string? where = rest.Count > 2 ? rest[2].ToLowerInvariant() switch
                        {
                            "centre" or "center" => "center",
                            "left" => "left",
                            "right" => "right",
                            "general" or "none" => "general",
                            _ => null,
                        } : null;
                        bool? wrap = Flag(args, "--wrap") ? true : Flag(args, "--nowrap") ? false : null;
                        if (where is null && wrap is null) { err.WriteLine("left, centre, right or general, and --wrap or --nowrap"); return 64; }
                        sheet.SetAlignment(cells, where, wrap);
                    }
                    else
                    {
                        var fill = Value(args, "--fill");
                        var ink = Value(args, "--ink");
                        if (fill is null && ink is null) { err.WriteLine("--fill RRGGBB and, or, --ink RRGGBB"); return 64; }
                        sheet.SetColours(cells, fill is "none" ? "" : fill, ink is "none" ? "" : ink);
                    }

                    file.Flush();
                    File.WriteAllBytes(file.Path, file.Package.ToBytes());
                    o.WriteLine($"{cells.Count} cell{(cells.Count == 1 ? "" : "s")} on {sheet.Name} changed");
                    return 0;
                }

                case "band":
                {
                    if (rest.Count < 1) { err.WriteLine("band <file> [--set <n> \"<text>\"]"); return 64; }
                    var file = PlainFile.Open(rest[0]);
                    var bands = HeaderFooter.All(file);
                    if (bands.Count == 0) { o.WriteLine("This document has no headers or footers."); return 1; }

                    for (int i = 0; i < bands.Count; i++)
                        o.WriteLine($"{i + 1}  {(bands[i].IsHeader ? "header" : "footer")}, {bands[i].Which}: "
                                  + bands[i].Text.Replace(Environment.NewLine, " / "));

                    var which = Value(args, "--set");
                    if (which is null) return 0;
                    if (!int.TryParse(which, out var index) || index < 1 || index > bands.Count)
                    { err.WriteLine($"there is no band {which}."); return 64; }
                    if (rest.Count < 2) { err.WriteLine("say what it should read: band <file> --set 1 \"the words\""); return 64; }

                    HeaderFooter.Write(file.Package, bands[index - 1].Part, rest[1]);
                    file.Flush();
                    File.WriteAllBytes(file.Path, file.Package.ToBytes());
                    o.WriteLine();
                    o.WriteLine($"band {index} now reads: {rest[1]}");
                    return 0;
                }

                case "links":
                {
                    if (rest.Count < 1) { err.WriteLine("links <file>"); return 64; }
                    var links = Links.All(PlainFile.Open(rest[0]));
                    if (links.Count == 0) { o.WriteLine("No links."); return 1; }
                    foreach (var link in links)
                    {
                        o.WriteLine($"{link.Where}: {link.Text}");
                        o.WriteLine($"    -> {(link.Target.Length == 0 ? "nowhere" : link.Target)}"
                                  + (link.External && !Links.SafeToOpen(link.Target) ? "   (Plain would not open this)" : ""));
                    }
                    return 0;
                }

                case "compare":
                {
                    if (rest.Count < 2) { err.WriteLine("compare <before> <after>"); return 64; }
                    var report = Compare.Between(PlainFile.Open(rest[0]), PlainFile.Open(rest[1]));

                    foreach (var part in report.Parts) o.WriteLine($"{part.How,-8} {part.Name}");
                    if (report.Parts.Count > 0) o.WriteLine();
                    foreach (var change in report.Text)
                        o.WriteLine($"{(change.Added ? "+" : "-")} {change.Where}: {change.Text}");

                    if (report.Note is not null) { o.WriteLine(); o.WriteLine(report.Note); }
                    o.WriteLine();
                    o.WriteLine($"{report.Parts.Count} part{(report.Parts.Count == 1 ? "" : "s")} differ, "
                              + $"{report.Same} the same, {report.Text.Count} line{(report.Text.Count == 1 ? "" : "s")} of text.");
                    return report.Parts.Count == 0 && report.Text.Count == 0 ? 0 : 1;
                }

                case "border":
                {
                    if (rest.Count < 3) { err.WriteLine("border <file> <range> <weight|none> [--sides ...] [--colour RRGGBB]"); return 64; }
                    var file = PlainFile.Open(rest[0]);
                    if (file.Workbook is null) { err.WriteLine("border only works on a spreadsheet."); return 2; }
                    var sheet = PickSheet(file.Workbook, Value(args, "--sheet"));
                    if (sheet is null) { err.WriteLine("no sheet by that name."); return 2; }

                    var cells = Block(rest[1]);
                    if (cells.Count == 0) { err.WriteLine("a range looks like B2 or A2:D40."); return 64; }

                    string weight = rest[2].Equals("none", StringComparison.OrdinalIgnoreCase) ? "" : rest[2].ToLowerInvariant();
                    if (weight.Length > 0 && !Styles.Weights.Any(w => w.Code == weight))
                    { err.WriteLine("weights: " + string.Join(", ", Styles.Weights.Select(w => w.Code)) + ", or none"); return 64; }

                    var sides = (Value(args, "--sides") ?? "all")
                        .Split(',', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries);
                    sheet.SetBorder(cells, sides, weight, Value(args, "--colour") ?? "");
                    file.Flush();
                    File.WriteAllBytes(file.Path, file.Package.ToBytes());
                    o.WriteLine($"{cells.Count} cell{(cells.Count == 1 ? "" : "s")} on {sheet.Name} changed");
                    return 0;
                }

                case "tidy":
                {
                    if (rest.Count < 3) { err.WriteLine("tidy <file> dedupe|split <range> [--on \",\"]"); return 64; }
                    var file = PlainFile.Open(rest[0]);
                    if (file.Workbook is null) { err.WriteLine("tidy only works on a spreadsheet."); return 2; }
                    var sheet = PickSheet(file.Workbook, Value(args, "--sheet"));
                    if (sheet is null) { err.WriteLine("no sheet by that name."); return 2; }

                    var ends = rest[2].Split(':');
                    if (!CellRef.TryParse(ends[0], out var from)) { err.WriteLine("a range looks like A2:D40."); return 64; }
                    var to = ends.Length > 1 && CellRef.TryParse(ends[1], out var second) ? second : from;

                    var outcome = rest[1].ToLowerInvariant() switch
                    {
                        "dedupe" => Tidy.RemoveDuplicates(file.Workbook, sheet,
                            Math.Min(from.Row, to.Row), Math.Max(from.Row, to.Row),
                            Math.Min(from.Column, to.Column), Math.Max(from.Column, to.Column)),
                        "split" => Tidy.SplitColumn(file.Workbook, sheet, from.Column,
                            Math.Min(from.Row, to.Row), Math.Max(from.Row, to.Row), Value(args, "--on") ?? ","),
                        _ => new Tidy.Refused("dedupe or split"),
                    };
                    if (outcome is Tidy.Refused refused) { err.WriteLine(refused.Reason); return 2; }
                    file.Flush();
                    File.WriteAllBytes(file.Path, file.Package.ToBytes());
                    o.WriteLine(((Tidy.Done)outcome).What);
                    return 0;
                }

                case "reads":
                {
                    if (rest.Count < 2) { err.WriteLine("reads <file> <cell> [--sheet <n>]"); return 64; }
                    var file = PlainFile.Open(rest[0]);
                    if (file.Workbook is null) { err.WriteLine("reads only works on a spreadsheet."); return 2; }
                    var sheet = PickSheet(file.Workbook, Value(args, "--sheet"));
                    if (sheet is null) { err.WriteLine("no sheet by that name."); return 2; }
                    if (!CellRef.TryParse(rest[1], out var cell)) { err.WriteLine("a cell is a letter and a number."); return 64; }

                    var reads = Traces.Reads(file.Workbook, sheet, cell);
                    var readBy = Traces.ReadBy(file.Workbook, sheet, cell);

                    o.WriteLine(reads.Count == 0 ? $"{cell} holds no formula, so it reads nothing."
                                                 : $"{cell} reads:");
                    foreach (var t in reads) o.WriteLine($"    {t.Where}   {t.What}");
                    o.WriteLine();
                    o.WriteLine(readBy.Count == 0 ? $"Nothing reads {cell}." : $"{cell} is read by:");
                    foreach (var t in readBy) o.WriteLine($"    {t.Where}   {t.What}");
                    return 0;
                }

                case "choices":
                {
                    if (rest.Count < 1) { err.WriteLine("choices <file> [--sheet <n>]"); return 64; }
                    var file = PlainFile.Open(rest[0]);
                    if (file.Workbook is null) { err.WriteLine("choices only works on a spreadsheet."); return 2; }

                    int found = 0;
                    foreach (var sheet in file.Workbook.Sheets)
                        foreach (var rule in Choices.On(file.Package, sheet))
                        {
                            found++;
                            o.WriteLine($"{sheet.Name}!{rule.Where}: {rule.Kind}");
                            if (rule.Allowed.Count > 0) o.WriteLine("    " + string.Join(", ", rule.Allowed));
                            else if (rule.Formula.Length > 0) o.WriteLine("    " + rule.Formula);
                            if (rule.Says.Length > 0) o.WriteLine("    " + rule.Says);
                        }
                    if (found == 0) { o.WriteLine("No cells in this workbook restrict what they will take."); return 1; }
                    return 0;
                }

                case "notes":
                {
                    if (rest.Count < 1) { err.WriteLine("notes <file> [--set <slide> \"<text>\"]"); return 64; }
                    var file = PlainFile.Open(rest[0]);
                    if (file.Deck is null) { err.WriteLine("notes only works on a presentation."); return 2; }

                    var notes = Described.Notes(file);
                    if (notes.Count == 0) { o.WriteLine("No slide in this deck has notes."); return 1; }
                    foreach (var note in notes)
                        o.WriteLine($"slide {note.Slide}: {note.Text.Replace(Environment.NewLine, " / ")}");

                    var which = Value(args, "--set");
                    if (which is null) return 0;
                    if (!int.TryParse(which, out var slide)) { err.WriteLine("--set takes a slide number."); return 64; }
                    var target = notes.FirstOrDefault(n => n.Slide == slide);
                    if (target is null) { err.WriteLine($"slide {slide} has no notes part to write into."); return 2; }
                    if (rest.Count < 2) { err.WriteLine("say what the notes should read."); return 64; }

                    Described.WriteNotes(file.Package, target.Part, rest[1]);
                    file.Flush();
                    File.WriteAllBytes(file.Path, file.Package.ToBytes());
                    o.WriteLine();
                    o.WriteLine($"slide {slide} now reads: {rest[1]}");
                    return 0;
                }

                case "describe":
                {
                    if (rest.Count < 1) { err.WriteLine("describe <file> [--set \"<text>\"]"); return 64; }
                    var file = PlainFile.Open(rest[0]);
                    var pictures = Described.Pictures(file);
                    if (pictures.Count == 0) { o.WriteLine("No pictures in this file."); return 1; }

                    foreach (var picture in pictures)
                        o.WriteLine($"{picture.Where}: {picture.Name}"
                                  + (picture.Text.Length == 0 ? "   (nothing describes it)" : $"   {picture.Text}"));

                    var text = Value(args, "--set");
                    if (text is null) return pictures.Any(x => x.Text.Length == 0) ? 1 : 0;

                    int changed = 0;
                    foreach (var picture in pictures.Where(x => x.Text.Length == 0))
                        changed += Described.DescribePictures(file, picture.Where, picture.Name, text);
                    file.Flush();
                    File.WriteAllBytes(file.Path, file.Package.ToBytes());
                    o.WriteLine();
                    o.WriteLine($"described {changed} that had nothing.");
                    return 0;
                }

                case "link":
                {
                    if (rest.Count < 3) { err.WriteLine("link <file> <paragraph> <address>|off"); return 64; }
                    var file = PlainFile.Open(rest[0]);
                    if (!int.TryParse(rest[1], out var block) || block < 1) { err.WriteLine("a paragraph is a number from 1."); return 64; }

                    var outcome = rest[2].Equals("off", StringComparison.OrdinalIgnoreCase)
                        ? Insert.Unlink(file, block - 1)
                        : Insert.Link(file, block - 1, rest[2]);
                    if (outcome is Insert.Refused refused) { err.WriteLine(refused.Reason); return 2; }
                    File.WriteAllBytes(file.Path, file.Package.ToBytes());
                    o.WriteLine(((Insert.Done)outcome).What);
                    return 0;
                }

                case "picture":
                {
                    if (rest.Count < 2) { err.WriteLine("picture <file> <image> [--cm 8]"); return 64; }
                    var file = PlainFile.Open(rest[0]);
                    double cm = double.TryParse(Value(args, "--cm"), System.Globalization.NumberStyles.Float,
                        System.Globalization.CultureInfo.InvariantCulture, out var given) ? given : 8;

                    var outcome = Insert.Picture(file, rest[1], cm);
                    if (outcome is Insert.Refused refused) { err.WriteLine(refused.Reason); return 2; }
                    File.WriteAllBytes(file.Path, file.Package.ToBytes());
                    o.WriteLine(((Insert.Done)outcome).What);
                    return 0;
                }

                case "explorer":
                {
                    string what = rest.Count > 0 ? rest[0].ToLowerInvariant() : "status";
                    switch (what)
                    {
                        case "on":  o.WriteLine(Explorer.Register()); return Explorer.RegisteredAnywhere() ? 0 : 2;
                        case "off": o.WriteLine(Explorer.Undo()); return Explorer.RegisteredAnywhere() ? 2 : 0;
                        case "status":
                            bool there = Explorer.RegisteredAnywhere();
                            o.WriteLine(there
                                ? (Explorer.Registered()
                                    ? "on, pointing at this copy of Plain"
                                    : "on, but pointing at a different copy of Plain")
                                : "off");
                            return there ? 0 : 1;
                        default: err.WriteLine("explorer on|off|status"); return 64;
                    }
                }

                case "sort":
                {
                    if (rest.Count < 3) { err.WriteLine("sort <file> <range like A2:D40> <col> [down] [sheet]"); return 64; }
                    var file = PlainFile.Open(rest[0]);
                    if (file.Workbook is null) { err.WriteLine("sort only works on a spreadsheet."); return 2; }

                    var parts = rest[1].Split(':');
                    if (parts.Length != 2 || !CellRef.TryParse(parts[0], out var from) || !CellRef.TryParse(parts[1], out var to))
                    { err.WriteLine("a range looks like A2:D40."); return 64; }
                    if (!CellRef.TryParse(rest[2] + "1", out var keyCell))
                    { err.WriteLine("a column is a letter like A or AB."); return 64; }

                    bool down = rest.Count > 3 && rest[3].Equals("down", StringComparison.OrdinalIgnoreCase);
                    string? sheetName = rest.Count > 4 ? rest[4] : (rest.Count > 3 && !down ? rest[3] : null);
                    var sheet = sheetName is not null
                        ? file.Workbook.Sheets.FirstOrDefault(x => x.Name.Equals(sheetName, StringComparison.OrdinalIgnoreCase))
                        : file.Workbook.Sheets[0];
                    if (sheet is null) { err.WriteLine($"no sheet called \"{sheetName}\"."); return 2; }

                    var outcome = Sort.Rows(file.Workbook, sheet,
                        Math.Min(from.Row, to.Row), Math.Max(from.Row, to.Row),
                        Math.Min(from.Column, to.Column), Math.Max(from.Column, to.Column),
                        keyCell.Column, !down);

                    if (outcome is Sort.Refused refused) { err.WriteLine(refused.Reason); return 2; }
                    var done = (Sort.Sorted)outcome;
                    file.Flush();
                    File.WriteAllBytes(file.Path, file.Package.ToBytes());
                    o.WriteLine($"sorted {sheet.Name} by column {rest[2].ToUpperInvariant()}, {done.RowsMoved} rows moved");
                    return 0;
                }

                case "width":
                {
                    if (rest.Count < 3) { err.WriteLine("width <file> <col> <chars|fit> [sheet]"); return 64; }
                    var file = PlainFile.Open(rest[0]);
                    if (file.Workbook is null) { err.WriteLine("width only works on a spreadsheet."); return 2; }
                    var sheet = rest.Count > 3
                        ? file.Workbook.Sheets.FirstOrDefault(x => x.Name.Equals(rest[3], StringComparison.OrdinalIgnoreCase))
                        : file.Workbook.Sheets[0];
                    if (sheet is null) { err.WriteLine($"no sheet called \"{rest[3]}\"."); return 2; }
                    if (!CellRef.TryParse(rest[1] + "1", out var asCell)) { err.WriteLine("a column is a letter like A or AB."); return 64; }

                    double want;
                    if (rest[2].Equals("fit", StringComparison.OrdinalIgnoreCase)) want = sheet.WidestChars(asCell.Column);
                    else if (!double.TryParse(rest[2], System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture, out want))
                    { err.WriteLine("a width is a number of characters, or the word fit."); return 64; }

                    sheet.SetWidthChars(asCell.Column, want);
                    file.Flush();
                    File.WriteAllBytes(file.Path, file.Package.ToBytes());
                    o.WriteLine($"column {rest[1].ToUpperInvariant()} on {sheet.Name} is now {sheet.WidthChars(asCell.Column)} characters wide");
                    return 0;
                }

                case "freeze":
                {
                    if (rest.Count < 2) { err.WriteLine("freeze <file> <rows> [sheet]"); return 64; }
                    var file = PlainFile.Open(rest[0]);
                    if (file.Workbook is null) { err.WriteLine("freeze only works on a spreadsheet."); return 2; }
                    var sheet = rest.Count > 2
                        ? file.Workbook.Sheets.FirstOrDefault(x => x.Name.Equals(rest[2], StringComparison.OrdinalIgnoreCase))
                        : file.Workbook.Sheets[0];
                    if (sheet is null) { err.WriteLine($"no sheet called \"{rest[2]}\"."); return 2; }
                    if (!int.TryParse(rest[1], out var rows) || rows < 0) { err.WriteLine("a number of rows, 0 for none."); return 64; }

                    sheet.SetFrozenRows(rows);
                    file.Flush();
                    File.WriteAllBytes(file.Path, file.Package.ToBytes());
                    o.WriteLine(rows == 0
                        ? $"{sheet.Name} scrolls freely again"
                        : $"the first {rows} row{(rows == 1 ? "" : "s")} of {sheet.Name} stay on screen");
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
                    var page = PageSetup.Default with
                    {
                        Paper = Value(args, "--paper") ?? "A4",
                        Landscape = Flag(args, "--landscape"),
                        MarginMm = double.TryParse(Value(args, "--margin"), System.Globalization.NumberStyles.Float,
                            System.Globalization.CultureInfo.InvariantCulture, out var mm) ? mm : 20,
                        Header = Value(args, "--header") ?? "",
                        Footer = Value(args, "--footer") ?? "Page {page} of {pages}",
                    };
                    var result = PdfExport.Build(file, Path.GetFileNameWithoutExtension(file.Path), page);
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
