namespace Plain.Core;

/// <summary>How much is in a file. A word count is on every essay's title page and in every writer's head.</summary>
public sealed record Tally(int Words, int Characters, int CharactersWithoutSpaces, int Paragraphs)
{
    public override string ToString() =>
        $"{Words:N0} words, {Characters:N0} characters, {Paragraphs:N0} paragraphs";
}

public static class Counts
{
    /// <summary>Count the words the way a person would: runs of anything that is not a space.</summary>
    public static Tally Of(string text)
    {
        int words = 0;
        bool inWord = false;
        int characters = 0, withoutSpaces = 0;
        foreach (char c in text)
        {
            if (c is '\n' or '\r') { inWord = false; continue; }
            characters++;
            if (char.IsWhiteSpace(c)) { inWord = false; continue; }
            withoutSpaces++;
            if (!inWord) { words++; inWord = true; }
        }
        int paragraphs = text.Split('\n').Count(line => line.Trim().Length > 0);
        return new Tally(words, characters, withoutSpaces, paragraphs);
    }

    public static Tally Of(PlainFile file) => file.Kind switch
    {
        FileKind.Document => Of(file.Document!.PlainText()),
        FileKind.Presentation => Of(string.Join("\n", file.Deck!.Slides
            .SelectMany(s => s.Texts()).SelectMany(t => t.Lines))),
        FileKind.Spreadsheet => Of(string.Join("\n", file.Workbook!.Sheets
            .SelectMany(s => s.Cells()).Select(c => c.Display))),
        _ => new Tally(0, 0, 0, 0),
    };

    /// <summary>The words in one part of a document, for a count of what is selected.</summary>
    public static Tally OfBlocks(Document doc, IEnumerable<int> blocks) =>
        Of(string.Join("\n", blocks.Select(i => doc.Read(i).Text)));
}
