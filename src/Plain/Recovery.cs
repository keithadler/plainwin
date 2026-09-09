using System.IO;
using System.Text.Json;
using Plain.Core;

namespace Plain;

/// <summary>
/// A copy of work that has not been saved yet, kept beside the settings so a machine that dies does not take the
/// afternoon with it. The power really does go three times a day in some of the places this runs.
///
/// It is a whole file, not a patch, written under a name that records where it came from. When Plain next starts it
/// offers any it finds; once the real file is saved the copy goes.
/// </summary>
public static class Recovery
{
    private sealed record Note(string Original, string When);

    private static string Key(string path) =>
        Convert.ToHexString(System.Security.Cryptography.SHA256.HashData(
            System.Text.Encoding.UTF8.GetBytes(path.ToLowerInvariant())))[..16];

    public static void Keep(PlainFile file, string path)
    {
        try
        {
            Directory.CreateDirectory(Settings.RecoveryFolder);
            var key = Key(path);
            File.WriteAllBytes(Path.Combine(Settings.RecoveryFolder, key + Path.GetExtension(path)), file.Package.ToBytes());
            File.WriteAllText(Path.Combine(Settings.RecoveryFolder, key + ".json"),
                JsonSerializer.Serialize(new Note(path, DateTimeOffset.Now.ToString("s"))));
        }
        catch { }   // a recovery copy that will not write is not worth interrupting anyone over
    }

    public static void Forget(string path)
    {
        try
        {
            var key = Key(path);
            foreach (var file in Directory.EnumerateFiles(Settings.RecoveryFolder, key + ".*")) File.Delete(file);
        }
        catch { }
    }

    /// <summary>What is waiting to be recovered: the file it came from, when it was kept, and the copy itself.</summary>
    public static IReadOnlyList<(string Original, string When, string Copy)> Waiting()
    {
        var found = new List<(string, string, string)>();
        try
        {
            if (!Directory.Exists(Settings.RecoveryFolder)) return found;
            foreach (var note in Directory.EnumerateFiles(Settings.RecoveryFolder, "*.json"))
            {
                var read = JsonSerializer.Deserialize<Note>(File.ReadAllText(note));
                if (read is null) continue;
                var copy = Directory.EnumerateFiles(Settings.RecoveryFolder,
                    Path.GetFileNameWithoutExtension(note) + ".*").FirstOrDefault(f => !f.EndsWith(".json"));
                if (copy is null) continue;
                found.Add((read.Original, read.When, copy));
            }
        }
        catch { }
        return found;
    }

    public static void ForgetAll()
    {
        try { if (Directory.Exists(Settings.RecoveryFolder)) Directory.Delete(Settings.RecoveryFolder, true); } catch { }
    }
}
