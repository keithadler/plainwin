namespace Plain.Core;

/// <summary>
/// What happened when a row was added to or taken out of a table. The work itself is on <see cref="Document"/>,
/// where the numbering of paragraphs lives and has to be worked out again afterwards.
/// </summary>
public static class TableRows
{
    public abstract record Result;
    public sealed record Done(string What) : Result;
    public sealed record Refused(string Reason) : Result;
}
