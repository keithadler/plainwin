namespace Plain;

/// <summary>
/// What a view has to do to take part in Find. Collecting the matches and revealing one are separate, so the window
/// can count them, step through them and say where you are without knowing anything about grids or slides.
/// </summary>
public interface IFindable
{
    /// <summary>Collect every match for a term and return how many there are.</summary>
    int FindAll(string term);

    /// <summary>Show the match at that position, and say where it is in words.</summary>
    string Reveal(int index);
}
