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

/// <summary>One line of the preserved list: what the part is, in words, and how big it is.</summary>
public sealed record PartNote(string Name, PartRole Role, string What, long Bytes)
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
            notes.Add(new PartNote(p.Name, role, What(p.Name), p.Size));
        }
        return notes;
    }

    /// <summary>What a part is, said the way someone who did not write the spec would say it.</summary>
    public static string What(string name)
    {
        string n = name.ToLowerInvariant();
        string file = n[(n.LastIndexOf('/') + 1)..];

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
        if (n.Contains("/charts/") || file.StartsWith("chart")) return "A chart";
        if (n.Contains("/diagrams/")) return "A SmartArt diagram";
        if (n.Contains("/drawings/") || file.StartsWith("drawing")) return "Drawing placement";
        if (n.Contains("theme")) return "The colour and font theme";

        // Spreadsheet
        if (n.Contains("pivotcache")) return "A pivot table's stored data";
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
        return "Part of the file Plain does not model";
    }
}
