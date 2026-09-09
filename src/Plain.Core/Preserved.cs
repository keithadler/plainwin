namespace Plain.Core;

public enum PartRole
{
    /// <summary>Plain reads and can edit this part.</summary>
    Shown,
    /// <summary>Plain changed this part in this session.</summary>
    Edited,
    /// <summary>Plain does not model this part. It is kept exactly as found.</summary>
    Preserved,
}

/// <summary>
/// Whether a part is worth a line of someone's attention. A chart or a macro is something you would want to know is
/// still in your file. The list of which part is which type is not: it is bookkeeping that every Office file carries
/// and nobody has ever wanted to see. Both are preserved exactly the same way; only the telling differs.
/// </summary>
public enum PartClass { Content, Bookkeeping }

/// <summary>One line of the preserved list: what the part is, in words, and how big it is.</summary>
public sealed record PartNote(string Name, PartRole Role, string What, long Bytes, PartClass Class = PartClass.Content)
{
    public string Size => Bytes >= 1_000_000 ? $"{Bytes / 1_048_576.0:0.0} MB"
                        : Bytes >= 1_000 ? $"{Bytes / 1024.0:0.0} KB"
                        : $"{Bytes} B";

    /// <summary>What a screen reader announces for this row.</summary>
    public override string ToString() => $"{What}, {Size}, {Name}";
}

/// <summary>
/// The list behind Plain's promise. Every part of the file is named and placed in one of three buckets, so what the
/// app cannot draw is still something you can see it holding on to.
/// </summary>
public static class Preserved
{
    public static IReadOnlyList<PartNote> Describe(OpcPackage pkg, IEnumerable<string> shownParts)
    {
        var shown = new HashSet<string>(shownParts.Select(OpcPackage.Normalize), StringComparer.Ordinal);
        var notes = new List<PartNote>();
        foreach (var p in pkg.Parts)
        {
            var role = p.Edited ? PartRole.Edited : shown.Contains(p.Name) ? PartRole.Shown : PartRole.Preserved;
            notes.Add(new PartNote(p.Name, role, What(p.Name), p.Size, Classify(p.Name)));
        }
        return notes;
    }

    /// <summary>One line of the panel: a kind of thing, how many of them, and how much room they take.</summary>
    public sealed record PreservedRow(string What, int Count, long Bytes, string? Name)
    {
        public string Size => Bytes >= 1_000_000 ? $"{Bytes / 1_048_576.0:0.0} MB"
                            : Bytes >= 1_000 ? $"{Bytes / 1024.0:0.0} KB"
                            : $"{Bytes} B";

        /// <summary>What the row reads as: the part's own name when there is one of it, a count when there are more.</summary>
        public string Title => Count == 1 ? What : $"{Count} {Many(What)}";

        public override string ToString() => $"{Title}, {Size}" + (Name is null ? "" : $", {Name}");
    }

    /// <summary>
    /// Gather the parts into lines worth reading. A document with two dozen embedded fonts should say so once, not
    /// fill the panel with two dozen indistinguishable rows, and bookkeeping should not take up a line at all.
    /// </summary>
    public static IReadOnlyList<PreservedRow> Summarise(IEnumerable<PartNote> notes)
    {
        var content = notes.Where(n => n.Role == PartRole.Preserved && n.Class == PartClass.Content).ToList();
        return content
            .GroupBy(n => n.What)
            .Select(g => new PreservedRow(g.Key, g.Count(), g.Sum(n => n.Bytes),
                                          g.Count() == 1 ? g.First().Name : null))
            .OrderByDescending(r => r.Bytes)
            .ToList();
    }

    /// <summary>How many parts are pure bookkeeping, which the panel mentions in one line rather than listing.</summary>
    public static (int Count, long Bytes) Bookkeeping(IEnumerable<PartNote> notes)
    {
        var kept = notes.Where(n => n.Role == PartRole.Preserved && n.Class == PartClass.Bookkeeping).ToList();
        return (kept.Count, kept.Sum(n => n.Bytes));
    }

    /// <summary>
    /// The plural of one of the phrases above. Every phrase either starts with "A"/"An" and ends in its own noun, or
    /// is already plural and left alone. That is a rule about the phrases as much as about English, so the tests keep
    /// a table of them: write "A pivot table's stored data" and this would cheerfully say "stored datas".
    /// </summary>
    public static string Many(string what)
    {
        string noun = what.StartsWith("An ", StringComparison.Ordinal) ? what[3..]
                    : what.StartsWith("A ", StringComparison.Ordinal) ? what[2..]
                    : null!;
        if (noun is null) return what;   // already plural, or a mass noun like "Macros"

        if (noun.EndsWith("y", StringComparison.Ordinal) && noun.Length > 1 && !"aeiou".Contains(noun[^2]))
            return noun[..^1] + "ies";
        if (noun.EndsWith("s", StringComparison.Ordinal) || noun.EndsWith("x", StringComparison.Ordinal)
            || noun.EndsWith("ch", StringComparison.Ordinal) || noun.EndsWith("sh", StringComparison.Ordinal))
            return noun + "es";
        return noun + "s";
    }

    /// <summary>
    /// Bookkeeping is everything a file needs to be a file, rather than anything someone put in it: which part is
    /// which, what points at what, the fonts a program listed, the settings it saved, the theme it started from.
    /// </summary>
    public static PartClass Classify(string name)
    {
        string n = name.ToLowerInvariant();
        string file = n[(n.LastIndexOf('/') + 1)..];

        if (n.EndsWith('/')) return PartClass.Bookkeeping;   // a folder marker, not a part with anything in it
        if (n.EndsWith(".rels")) return PartClass.Bookkeeping;
        if (n == "[content_types].xml") return PartClass.Bookkeeping;
        if (n.StartsWith("docprops/")) return PartClass.Bookkeeping;
        if (n.Contains("customxml")) return PartClass.Bookkeeping;
        if (file is "settings.xml" or "websettings.xml" or "fonttable.xml" or "styles.xml"
                 or "numbering.xml" or "calcchain.xml" or "presprops.xml" or "viewprops.xml"
                 or "tablestyles.xml" or "stylessheet.xml") return PartClass.Bookkeeping;
        if (n.Contains("theme")) return PartClass.Bookkeeping;
        if (n.Contains("/printersettings/")) return PartClass.Bookkeeping;
        return PartClass.Content;
    }

    /// <summary>What a part is, said the way someone who did not write the spec would say it.</summary>
    public static string What(string name)
    {
        string n = name.ToLowerInvariant();
        string file = n[(n.LastIndexOf('/') + 1)..];

        if (n.EndsWith('/')) return "A folder marker";
        if (n.EndsWith(".rels")) return "Links between parts";
        if (n == "[content_types].xml") return "The list of what each part is";
        if (n.StartsWith("docprops/")) return file switch
        {
            "core.xml" => "Author, title and dates",
            "app.xml" => "Word and page counts",
            "custom.xml" => "Custom properties",
            _ => "Document properties",
        };
        if (n.Contains("/media/") || n.Contains("/images/")) return "A picture";
        if (n.Contains("/embeddings/")) return "An embedded file";
        if (n.Contains("vbaproject.bin")) return "Macros";
        if (n.Contains("/fonts/") || file.EndsWith(".odttf") || file.EndsWith(".ttf")) return "An embedded font";
        if (file == "websettings.xml") return "Web settings";
        if (n.Contains("/printersettings/")) return "Printer settings";
        if (n.Contains("/charts/") || file.StartsWith("chart")) return "A chart";
        if (n.Contains("/diagrams/")) return "A SmartArt diagram";
        if (n.Contains("/drawings/") || file.StartsWith("drawing")) return "Drawing placement";
        if (n.Contains("theme")) return "The colour and font theme";

        // Spreadsheet
        if (n.Contains("pivotcache")) return "Pivot table data";
        if (n.Contains("pivottable")) return "A pivot table";
        if (file == "sharedstrings.xml") return "The text every sheet shares";
        if (file == "styles.xml") return "Number formats, fonts and fills";
        if (file == "calcchain.xml") return "The order Excel recalculates in";
        if (n.Contains("threadedcomment")) return "Comment threads";
        if (n.Contains("/tables/")) return "A named table";
        if (n.Contains("querytable")) return "An external data query";
        if (file == "workbook.xml") return "The list of sheets";
        if (n.Contains("/worksheets/")) return "A sheet";

        // Word
        if (file == "document.xml") return "The document text";
        if (file == "settings.xml") return "Document settings";
        if (file == "numbering.xml") return "List numbering";
        if (file == "fonttable.xml") return "The fonts used";
        if (file.StartsWith("header")) return "A page header";
        if (file.StartsWith("footer")) return "A page footer";
        if (file == "footnotes.xml") return "Footnotes";
        if (file == "endnotes.xml") return "Endnotes";
        if (file == "people.xml") return "Who made the tracked changes";
        if (file.StartsWith("comments")) return "Comments";
        if (n.Contains("/glossary/")) return "Reusable building blocks";

        // PowerPoint
        if (n.Contains("/slidelayouts/")) return "A slide layout";
        if (n.Contains("/slidemasters/")) return "A slide master";
        if (n.Contains("/notesslides/")) return "Speaker notes";
        if (n.Contains("/notesmasters/")) return "The notes master";
        if (n.Contains("/slides/")) return "A slide";
        if (file == "presentation.xml") return "The list of slides";
        if (n.Contains("presprops")) return "Presentation settings";
        if (n.Contains("tablestyles")) return "Table styles";
        if (n.Contains("viewprops")) return "Saved view settings";

        if (n.Contains("customxml")) return "Custom XML kept by another program";
        return "An unrecognised part";
    }
}
