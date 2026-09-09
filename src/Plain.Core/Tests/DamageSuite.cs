namespace Plain.Core.Tests;

/// <summary>
/// People open files that arrived by email, off a memory stick, out of a half-finished download. A damaged one must
/// produce a sentence someone can act on, never a crash and never a hang. This takes each fixture apart in the ways
/// files really break - a flipped byte, a truncated download, a zeroed block, a lost directory - and insists that
/// whatever comes back is either a working file or a clear refusal.
/// </summary>
public static class DamageSuite
{
    public static Suite Run()
    {
        var s = new Suite("damage");
        var names = new[] { "sheet.xlsx", "doc.docx", "deck.pptx" };
        if (Fixtures.Missing(s, names)) return s;
        var random = new Random(20260908);   // fixed, so a failure can be reproduced
        // Deeper runs on demand: PLAIN_FUZZ=5000 dotnet run --project src/Plain.Selftest -- damage
        int flips = int.TryParse(Environment.GetEnvironmentVariable("PLAIN_FUZZ"), out var n) && n > 0 ? n : 400;

        int opened = 0, refused = 0, leaked = 0;
        string firstLeak = "";

        foreach (var name in names)
        {
            var original = File.ReadAllBytes(Fixtures.Path_(name)!);

            foreach (var (how, damaged) in Damage(original, random, flips))
            {
                try
                {
                    var file = PlainFile.Read(damaged);
                    // If it opened, everything it offers must be readable without throwing something unexpected.
                    foreach (var part in file.Package.Parts) _ = file.Package.Read(part.Name);
                    _ = file.Parts();
                    switch (file.Kind)
                    {
                        case FileKind.Spreadsheet: foreach (var sheet in file.Workbook!.Sheets) _ = sheet.Cells().Count(); break;
                        case FileKind.Document: _ = file.Document!.PlainText(); break;
                        case FileKind.Presentation: foreach (var slide in file.Deck!.Slides) _ = slide.Title(); break;
                    }
                    opened++;
                }
                catch (OpcPackage.PackageException) { refused++; }
                catch (Exception ex)
                {
                    leaked++;
                    if (firstLeak.Length == 0) firstLeak = $"{name} after {how}: {ex.GetType().Name}: {ex.Message}";
                }
            }
        }

        // Both outcomes have to happen, or the suite is only testing one path. A flipped bit often lands somewhere
        // harmless and the file still opens; damage to the directory or to a part Plain reads must be refused.
        s.Check($"some damaged files still opened ({opened} of {opened + refused + leaked})", opened > 0);
        s.Check($"some damaged files were refused ({refused} of {opened + refused + leaked})", refused > 0);
        s.Equal($"no damaged file threw something unexpected ({firstLeak})", 0, leaked);
        // ---- files that are not damaged, but are not this kind of file either ----
        // Saying "no ZIP end record found" about a password-protected file sends someone looking for damage that
        // is not there. These check the message is about what the file actually is.
        {
            var ole = new byte[] { 0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1 };

            var oldFile = new byte[2048];
            Array.Copy(ole, oldFile, ole.Length);
            string oldSaid = "";
            try { OpcPackage.Read(oldFile); } catch (Exception ex) { oldSaid = ex.Message; }
            s.Check("an older .doc is named as an older file", oldSaid.Contains("older Office file"));
            s.Check("and it says what to do about it", oldSaid.Contains(".docx"));

            var locked = new byte[4096];
            Array.Copy(ole, locked, ole.Length);
            var name = System.Text.Encoding.Unicode.GetBytes("EncryptedPackage");
            Array.Copy(name, 0, locked, 600, name.Length);
            string lockedSaid = "";
            try { OpcPackage.Read(locked); } catch (Exception ex) { lockedSaid = ex.Message; }
            s.Check("a file with a password says so", lockedSaid.Contains("password"));
            s.Check("and does not pretend it can take it off", lockedSaid.Contains("cannot remove"));
            s.Check("and says what would work instead", lockedSaid.Contains("save a copy"));
        }

        return s;
    }

    /// <summary>The ways a file really arrives broken, plus a spread of single flipped bytes.</summary>
    private static IEnumerable<(string How, byte[] Bytes)> Damage(byte[] original, Random random, int flips)
    {
        yield return ("truncated to a tenth", original[..(original.Length / 10)]);
        yield return ("truncated to half", original[..(original.Length / 2)]);
        yield return ("missing its last 40 bytes", original[..Math.Max(0, original.Length - 40)]);
        yield return ("empty", Array.Empty<byte>());
        yield return ("one byte", new byte[] { 0x50 });

        var zeroedStart = (byte[])original.Clone();
        Array.Clear(zeroedStart, 0, Math.Min(200, zeroedStart.Length));
        yield return ("first 200 bytes zeroed", zeroedStart);

        var zeroedEnd = (byte[])original.Clone();
        Array.Clear(zeroedEnd, Math.Max(0, zeroedEnd.Length - 200), Math.Min(200, zeroedEnd.Length));
        yield return ("last 200 bytes zeroed", zeroedEnd);

        var zeroedMiddle = (byte[])original.Clone();
        Array.Clear(zeroedMiddle, zeroedMiddle.Length / 3, Math.Min(500, zeroedMiddle.Length / 3));
        yield return ("a block in the middle zeroed", zeroedMiddle);

        var garbage = (byte[])original.Clone();
        random.NextBytes(garbage.AsSpan(garbage.Length / 2, Math.Min(300, garbage.Length / 2)));
        yield return ("random bytes over the middle", garbage);

        for (int i = 0; i < flips; i++)
        {
            var flipped = (byte[])original.Clone();
            int at = random.Next(flipped.Length);
            flipped[at] ^= (byte)(1 << random.Next(8));
            yield return ($"a bit flipped at {at}", flipped);
        }
    }
}
