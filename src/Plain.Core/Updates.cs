using System.Text.Json;

namespace Plain.Core;

/// <summary>
/// Updates without a framework: one GET to GitHub's releases API, compare versions, and say so. Nothing is
/// downloaded and nothing installs itself; the most it ever does is open the release page in your browser when you
/// ask it to. No identifiers are sent, nothing about your files is sent, and the whole thing can be turned off.
/// The parts that decide anything live here, with no network in them, so they can be checked without one.
/// </summary>
public static class Updates
{
    public const string Repo = "keithadler/plainwin";
    public static readonly string ReleasesPage = $"https://github.com/{Repo}/releases/latest";
    public static readonly string Api = $"https://api.github.com/repos/{Repo}/releases/latest";
    public static readonly TimeSpan Interval = TimeSpan.FromHours(24);

    public abstract record Result;
    public sealed record UpToDate(string Latest) : Result;
    public sealed record Available(string Version, string Page) : Result;
    public sealed record Unknown(string Reason) : Result;

    /// <summary>What GitHub's answer means. A failure is never worth interrupting anyone over.</summary>
    public static Result Parse(int status, string body, string current)
    {
        if (status == 404) return new Unknown("No public release yet.");
        if (status != 200) return new Unknown($"HTTP {status}");
        try
        {
            using var doc = JsonDocument.Parse(body);
            if (!doc.RootElement.TryGetProperty("tag_name", out var tagElement))
                return new Unknown("No tag in the answer.");
            var tag = tagElement.GetString() ?? "";
            var latest = tag.StartsWith('v') ? tag[1..] : tag;
            if (latest.Length == 0) return new Unknown("No tag in the answer.");
            var page = doc.RootElement.TryGetProperty("html_url", out var u) ? u.GetString() ?? ReleasesPage : ReleasesPage;
            return IsNewer(latest, current) ? new Available(latest, page) : new UpToDate(latest);
        }
        catch (Exception ex) { return new Unknown(ex.Message); }
    }

    /// <summary>Compare two dotted versions a piece at a time, so 1.10.0 is after 1.9.0 rather than before it.</summary>
    public static bool IsNewer(string a, string b)
    {
        var pa = a.Split('.').Select(s => int.TryParse(s, out var n) ? n : 0).ToArray();
        var pb = b.Split('.').Select(s => int.TryParse(s, out var n) ? n : 0).ToArray();
        for (int i = 0; i < Math.Max(pa.Length, pb.Length); i++)
        {
            int x = i < pa.Length ? pa[i] : 0, y = i < pb.Length ? pb[i] : 0;
            if (x != y) return x > y;
        }
        return false;
    }

    /// <summary>
    /// Is a check due? An hour of slack, so opening the app a little earlier than yesterday still counts as a day
    /// and it does not quietly become an every-other-day check.
    /// </summary>
    public static bool ShouldCheck(bool enabled, DateTimeOffset? last, DateTimeOffset now)
        => enabled && (last is null || now - last.Value >= Interval - TimeSpan.FromHours(1));
}
