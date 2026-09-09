using System.Diagnostics;

namespace Plain.Core.Tests;

/// <summary>A named group of checks. Plain ships its tests inside the app so anyone can run them: "plain selftest".</summary>
public sealed class Suite
{
    public string Name { get; }
    public List<(string Name, bool Ok, string? Detail)> Results { get; } = new();
    public int Passed => Results.Count(r => r.Ok);
    public int Failed => Results.Count(r => !r.Ok);

    public Suite(string name) { Name = name; }

    public void Check(string name, bool ok, string? detail = null) => Results.Add((name, ok, ok ? null : detail));

    public void Equal<T>(string name, T expected, T actual)
    {
        bool ok = EqualityComparer<T>.Default.Equals(expected, actual);
        Results.Add((name, ok, ok ? null : $"expected {expected}, got {actual}"));
    }

    public void Bytes(string name, byte[] expected, byte[] actual)
    {
        if (expected.Length != actual.Length) { Results.Add((name, false, $"length {expected.Length} vs {actual.Length}")); return; }
        for (int i = 0; i < expected.Length; i++)
            if (expected[i] != actual[i]) { Results.Add((name, false, $"first difference at byte {i}: {expected[i]:x2} vs {actual[i]:x2}")); return; }
        Results.Add((name, true, null));
    }

    public void Throws<T>(string name, Action action) where T : Exception
    {
        try { action(); Results.Add((name, false, "no exception was thrown")); }
        catch (T) { Results.Add((name, true, null)); }
        catch (Exception ex) { Results.Add((name, false, $"threw {ex.GetType().Name}: {ex.Message}")); }
    }
}

public static class SelfTest
{
    private static readonly (string Name, Func<Suite> Run)[] All =
    {
        ("opc", OpcSuite.Run),
        ("refs", RefsSuite.Run),
        ("dialect", DialectSuite.Run),
        ("replace", ReplaceSuite.Run),
        ("grid", GridSuite.Run),
        ("formula", FormulaSuite.Run),
        ("annotations", AnnotationsSuite.Run),
        ("format", FormatSuite.Run),
        ("roundtwo", RoundTwoSuite.Run),
        ("sheet", SheetSuite.Run),
        ("doc", DocSuite.Run),
        ("deck", DeckSuite.Run),
        ("scale", ScaleSuite.Run),
        ("damage", DamageSuite.Run),
        ("sort", SortSuite.Run),
        ("page", PageSuite.Run),
        ("slides", SlidesSuite.Run),
        ("table", TableSuite.Run),
        ("hidden", HiddenSuite.Run),
        ("folder", FolderSuite.Run),
        ("sheets", SheetsSuite.Run),
        ("docparts", DocPartsSuite.Run),
        ("extras", ExtrasSuite.Run),
        ("updates", UpdatesSuite.Run),
        ("corpus", CorpusSuite.Run),
    };

    public static int Run(TextWriter o, string? only = null, bool list = false)
    {
        if (list) { foreach (var s in All) o.WriteLine(s.Name); return 0; }

        int passed = 0, failed = 0;
        var sw = Stopwatch.StartNew();
        foreach (var (name, run) in All)
        {
            if (only is not null && !string.Equals(only, name, StringComparison.OrdinalIgnoreCase)) continue;
            Suite suite;
            try { suite = run(); }
            catch (Exception ex) { o.WriteLine($"  FAIL  {name}: the suite itself threw {ex.GetType().Name}: {ex.Message}"); failed++; continue; }
            foreach (var r in suite.Results)
            {
                if (r.Ok) passed++;
                else { failed++; o.WriteLine($"  FAIL  {suite.Name}: {r.Name}{(r.Detail is null ? "" : " - " + r.Detail)}"); }
            }
        }
        o.WriteLine($"{passed} passed, {failed} failed in {sw.ElapsedMilliseconds} ms");
        return failed == 0 ? 0 : 1;
    }
}
