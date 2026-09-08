namespace Plain.Core.Tests;

/// <summary>Finds the fixture files whether the tests run from the repo, from bin, or from an installed exe.</summary>
public static class Fixtures
{
    public static string? Dir
    {
        get
        {
            var set = Environment.GetEnvironmentVariable("PLAIN_FIXTURES");
            if (!string.IsNullOrWhiteSpace(set) && Directory.Exists(set)) return set;
            var dir = new DirectoryInfo(AppContext.BaseDirectory);
            for (int i = 0; i < 8 && dir is not null; i++, dir = dir.Parent)
            {
                var candidate = Path.Combine(dir.FullName, "tests", "fixtures");
                if (Directory.Exists(candidate)) return candidate;
            }
            return null;
        }
    }

    public static string? Path_(string name)
    {
        var dir = Dir;
        if (dir is null) return null;
        var p = System.IO.Path.Combine(dir, name);
        return File.Exists(p) ? p : null;
    }

    /// <summary>
    /// True when any of these fixtures is not to hand, having noted it as a skip rather than a failure. Someone who
    /// downloads the exe and runs "plain selftest" has no fixtures; they should see the checks that can run pass,
    /// not a wall of errors about files they were never given.
    /// </summary>
    public static bool Missing(Suite suite, params string[] names)
    {
        var absent = names.Where(n => Path_(n) is null).ToList();
        if (absent.Count == 0) return false;
        suite.Check($"skipped: needs {string.Join(", ", absent)} from the repository's tests/fixtures", true);
        return true;
    }

    /// <summary>A scratch copy, so a test never writes over a fixture.</summary>
    public static string Copy(string name)
    {
        var src = Path_(name) ?? throw new FileNotFoundException($"fixture {name} not found");
        var dst = System.IO.Path.Combine(System.IO.Path.GetTempPath(), $"plain-{Guid.NewGuid():N}-{name}");
        File.Copy(src, dst, true);
        return dst;
    }
}
