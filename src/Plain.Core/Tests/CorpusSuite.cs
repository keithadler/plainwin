namespace Plain.Core.Tests;

/// <summary>
/// The test that matters: take real Office files, open them, save them without changing anything, and demand the
/// output file equal the input byte for byte. Point PLAIN_CORPUS at a folder of real .docx/.xlsx/.pptx files to run
/// it over them. Without that variable the suite says so and passes, because most people have no corpus to hand.
/// </summary>
public static class CorpusSuite
{
    public static Suite Run()
    {
        var s = new Suite("corpus");
        var dir = Environment.GetEnvironmentVariable("PLAIN_CORPUS");
        if (string.IsNullOrWhiteSpace(dir) || !Directory.Exists(dir))
        {
            s.Check("no corpus folder set (PLAIN_CORPUS); skipping the real-file round trip", true);
            return s;
        }

        var files = Directory.EnumerateFiles(dir, "*", SearchOption.AllDirectories)
            .Where(f => f.EndsWith(".docx", StringComparison.OrdinalIgnoreCase)
                     || f.EndsWith(".xlsx", StringComparison.OrdinalIgnoreCase)
                     || f.EndsWith(".pptx", StringComparison.OrdinalIgnoreCase))
            .Where(f => !Path.GetFileName(f).StartsWith("~$"))
            .OrderBy(f => f, StringComparer.Ordinal).ToList();

        s.Check($"corpus has files ({files.Count} found)", files.Count > 0);

        int identical = 0, readable = 0;
        foreach (var f in files)
        {
            byte[] original;
            try { original = File.ReadAllBytes(f); } catch { continue; }
            try
            {
                var pkg = OpcPackage.Read(original);
                readable++;
                // Every part must decompress without error before we trust the round trip.
                foreach (var part in pkg.Parts) _ = pkg.Read(part.Name);
                var saved = pkg.ToBytes();
                if (saved.Length == original.Length && saved.AsSpan().SequenceEqual(original)) identical++;
                else s.Check($"round trip is byte identical: {Path.GetFileName(f)}", false,
                             saved.Length != original.Length ? $"length {original.Length} vs {saved.Length}" : "content differs");
            }
            catch (Exception ex)
            {
                s.Check($"opens: {Path.GetFileName(f)}", false, $"{ex.GetType().Name}: {ex.Message}");
            }
        }

        s.Equal("every corpus file opened", files.Count, readable);
        s.Equal("every corpus file round tripped byte for byte", files.Count, identical);
        return s;
    }
}
