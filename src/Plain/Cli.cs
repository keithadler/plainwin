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
    public const string Version = "1.0.0";

    private static readonly string[] Verbs =
        { "info", "parts", "text", "cells", "get", "set", "roundtrip", "selftest", "version", "help", "--help", "-h", "--version" };

    public static bool IsVerb(string arg) => Verbs.Contains(arg, StringComparer.OrdinalIgnoreCase);

    private const string Usage = """
        Plain for Windows - opens Word, Excel and PowerPoint files, edits the basics,
        and never damages what it doesn't understand.

          plain info <file>                what the file is, and what Plain keeps untouched
          plain parts <file> [--json]      every part, and whether Plain shows it or preserves it
          plain text <file>                the text, as plain text
          plain cells <file> [sheet]       every filled cell, as reference<tab>value
          plain get <file> <ref>           one cell, or one block by number
          plain set <file> <ref> <value>   change one cell or block, then save in place
          plain roundtrip <file>...        prove a save changes nothing: byte compares the result
          plain selftest [suite]           run the built-in checks
          plain version

        A cell is A1. A block is a number from "plain text --numbered". A value that starts
        with = is a formula. Nothing is uploaded and nothing is sent anywhere, ever.
        Exit codes: 0 fine, 1 something to look at, 2 problem, 64 usage.
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
                    return 0;

                case "selftest":
                    return Core.Tests.SelfTest.Run(o, rest.FirstOrDefault(), args.Contains("--list"));

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

    private static string Describe(FileKind kind) => kind switch
    {
        FileKind.Spreadsheet => "Excel workbook",
        FileKind.Document => "Word document",
        FileKind.Presentation => "PowerPoint deck",
        _ => "unknown",
    };
}
