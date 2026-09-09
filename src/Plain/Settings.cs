using System.IO;
using System.Text.Json;
using System.Text.Json.Serialization;

namespace Plain;

/// <summary>
/// The few things Plain remembers between runs. It remembers them because people asked it to: a list of what you had
/// open, how big the text is, and whether you want the preserved panel showing. Nothing here is about you and none of
/// it leaves the machine; it is one small file you can delete at any time.
/// </summary>
public sealed class Settings
{
    public List<string> Recent { get; set; } = new();
    public double TextScale { get; set; } = 1.0;
    public bool ShowPreserved { get; set; } = true;

    [JsonIgnore]
    public static string Folder =>
        Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "Plain for Windows");

    [JsonIgnore]
    public static string FilePath => Path.Combine(Folder, "settings.json");

    private static readonly JsonSerializerOptions Format = new() { WriteIndented = true };

    public static Settings Load()
    {
        try
        {
            if (File.Exists(FilePath))
                return JsonSerializer.Deserialize<Settings>(File.ReadAllText(FilePath)) ?? new Settings();
        }
        catch { }
        return new Settings();
    }

    public void Save()
    {
        try
        {
            Directory.CreateDirectory(Folder);
            File.WriteAllText(FilePath, JsonSerializer.Serialize(this, Format));
        }
        catch { }   // a settings file that will not write is not worth interrupting anyone over
    }

    /// <summary>Put a file at the top of the list, keeping the ten most recent and dropping any that have gone.</summary>
    public void Remember(string path)
    {
        Recent.RemoveAll(p => string.Equals(p, path, StringComparison.OrdinalIgnoreCase));
        Recent.Insert(0, path);
        if (Recent.Count > 10) Recent.RemoveRange(10, Recent.Count - 10);
        Save();
    }

    /// <summary>The remembered files that are still there, so the list never offers something that has gone.</summary>
    public IReadOnlyList<string> RecentThatExist()
    {
        var live = Recent.Where(File.Exists).ToList();
        if (live.Count != Recent.Count) { Recent = live; Save(); }
        return live;
    }

    /// <summary>Text size steps, so Ctrl+plus and Ctrl+minus land somewhere sensible each time.</summary>
    public static readonly double[] Scales = { 1.0, 1.15, 1.3, 1.5, 1.75, 2.0, 2.5 };

    public double Bigger() => TextScale = Scales.FirstOrDefault(x => x > TextScale + 0.001, Scales[^1]);
    public double Smaller() => TextScale = Scales.LastOrDefault(x => x < TextScale - 0.001, Scales[0]);
}
