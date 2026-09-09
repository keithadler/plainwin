namespace Plain.Core.Tests;

/// <summary>
/// The update check, without a network. Everything that decides anything is a pure function, so what GitHub might
/// answer can be handed to it directly, including the answers that are not JSON at all.
/// </summary>
public static class UpdatesSuite
{
    public static Suite Run()
    {
        var s = new Suite("updates");

        s.Check("a later version is newer", Updates.IsNewer("1.0.2", "1.0.1"));
        s.Check("the same version is not newer", !Updates.IsNewer("1.0.1", "1.0.1"));
        s.Check("an earlier version is not newer", !Updates.IsNewer("1.0.0", "1.0.1"));
        s.Check("ten is after nine, not before it", Updates.IsNewer("1.10.0", "1.9.0"));
        s.Check("a new middle number counts", Updates.IsNewer("1.1.0", "1.0.9"));
        s.Check("a new first number counts", Updates.IsNewer("2.0.0", "1.99.99"));
        s.Check("a short version compares against a long one", Updates.IsNewer("1.1", "1.0.9"));
        s.Check("rubbish in a version is not newer", !Updates.IsNewer("banana", "1.0.1"));

        string body(string tag) => "{\"tag_name\":\"" + tag + "\",\"html_url\":\"https://example.invalid/r\"}";

        s.Check("a newer tag is offered",
            Updates.Parse(200, body("v1.0.2"), "1.0.1") is Updates.Available { Version: "1.0.2" });
        s.Check("the offer carries the page it came with",
            Updates.Parse(200, body("v1.0.2"), "1.0.1") is Updates.Available { Page: "https://example.invalid/r" });
        s.Check("the same tag is up to date",
            Updates.Parse(200, body("v1.0.1"), "1.0.1") is Updates.UpToDate);
        s.Check("an older tag is up to date",
            Updates.Parse(200, body("v1.0.0"), "1.0.1") is Updates.UpToDate);
        s.Check("a tag without the v still reads",
            Updates.Parse(200, body("1.0.2"), "1.0.1") is Updates.Available { Version: "1.0.2" });

        s.Check("no release yet is not an error to show",
            Updates.Parse(404, "", "1.0.1") is Updates.Unknown);
        s.Check("rate limiting is not an error to show",
            Updates.Parse(403, "", "1.0.1") is Updates.Unknown);
        s.Check("a server having a bad day says nothing",
            Updates.Parse(500, "", "1.0.1") is Updates.Unknown);
        s.Check("an answer that is not JSON says nothing",
            Updates.Parse(200, "<html>who knows</html>", "1.0.1") is Updates.Unknown);
        s.Check("JSON without a tag says nothing",
            Updates.Parse(200, "{\"nothing\":true}", "1.0.1") is Updates.Unknown);
        s.Check("an empty tag says nothing",
            Updates.Parse(200, body(""), "1.0.1") is Updates.Unknown);
        s.Check("an empty body says nothing",
            Updates.Parse(200, "", "1.0.1") is Updates.Unknown);

        var now = DateTimeOffset.Parse("2026-09-08T12:00:00Z");
        s.Check("turned off means never", !Updates.ShouldCheck(false, null, now));
        s.Check("turned off stays off however long it has been",
            !Updates.ShouldCheck(false, now - TimeSpan.FromDays(30), now));
        s.Check("never checked means check", Updates.ShouldCheck(true, null, now));
        s.Check("checked a minute ago means wait",
            !Updates.ShouldCheck(true, now - TimeSpan.FromMinutes(1), now));
        s.Check("checked a day ago means check",
            Updates.ShouldCheck(true, now - TimeSpan.FromHours(24), now));
        s.Check("opening an hour earlier than yesterday still counts as a day",
            Updates.ShouldCheck(true, now - TimeSpan.FromHours(23), now));
        s.Check("half a day is not a day",
            !Updates.ShouldCheck(true, now - TimeSpan.FromHours(12), now));
        s.Check("a clock that went backwards does not make it check every time",
            !Updates.ShouldCheck(true, now + TimeSpan.FromHours(2), now));

        s.Check("it asks GitHub and nowhere else", Updates.Api.StartsWith("https://api.github.com/"));
        s.Check("it asks about this repository", Updates.Api.Contains("keithadler/plainwin"));
        s.Check("the page it would open is this repository's",
            Updates.ReleasesPage == "https://github.com/keithadler/plainwin/releases/latest");

        return s;
    }
}
