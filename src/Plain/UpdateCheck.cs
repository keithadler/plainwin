using System.Net.Http;
using System.Windows;
using Plain.Core;

namespace Plain;

/// <summary>
/// The one place Plain touches a network. Once a day, in the background, it asks GitHub what the latest release is
/// and does nothing else with the answer than put a line in the status bar. It sends no identifier and nothing about
/// your files, it downloads nothing, it installs nothing, and it never interrupts: a failed check is silent, because
/// somebody working offline in a hotel does not need a dialog about it.
///
/// The switch is in Reading and settings, on by default like the rest of the family. Turned off, this class makes no
/// request at all and Plain reaches the network never.
/// </summary>
public static class UpdateCheck
{
    /// <summary>Set when a check found something, so the window can offer it. Null until then.</summary>
    public static Updates.Available? Found { get; private set; }

    private static readonly HttpClient Client = new() { Timeout = TimeSpan.FromSeconds(8) };

    /// <summary>
    /// Run a check if one is due, then hand the answer back on the UI thread. Nothing here throws into the caller:
    /// the whole point is that the app is unaffected by whether this works.
    /// </summary>
    public static void Start(Settings settings, Action<Updates.Available> onFound)
    {
        if (!Updates.ShouldCheck(settings.CheckForUpdates, settings.LastUpdateCheck, DateTimeOffset.UtcNow)) return;

        _ = Task.Run(async () =>
        {
            var result = await Ask(Cli.Version).ConfigureAwait(false);

            // Record the attempt whatever came back, so a machine with no network does not try on every keystroke.
            settings.LastUpdateCheck = DateTimeOffset.UtcNow;

            if (result is Updates.Available available && available.Version != settings.UpdateSeen)
            {
                settings.UpdateSeen = available.Version;
                Found = available;
                try
                {
                    Application.Current?.Dispatcher.BeginInvoke(() => onFound(available));
                }
                catch { /* the window went away first; nothing to tell */ }
            }

            try { settings.Save(); } catch { /* a settings file that will not write is not worth a message */ }
        });
    }

    /// <summary>One GET, with everything that can go wrong turned into "do not know".</summary>
    public static async Task<Updates.Result> Ask(string current)
    {
        try
        {
            using var request = new HttpRequestMessage(HttpMethod.Get, Updates.Api);
            // GitHub wants a user agent, and this is the only thing Plain ever says about itself.
            request.Headers.Add("User-Agent", "Plain-for-Windows/" + current);
            request.Headers.Add("Accept", "application/vnd.github+json");
            using var response = await Client.SendAsync(request).ConfigureAwait(false);
            var body = await response.Content.ReadAsStringAsync().ConfigureAwait(false);
            return Updates.Parse((int)response.StatusCode, body, current);
        }
        catch (Exception ex) { return new Updates.Unknown(ex.Message); }
    }

    /// <summary>
    /// Open an address in the browser, which is as far as Plain will go on your behalf. Used for the release page
    /// and for a link in a document that has been checked as an ordinary web or mail address.
    /// </summary>
    public static void OpenReleasePage(string page)
    {
        try
        {
            System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(page) { UseShellExecute = true });
        }
        catch { /* no browser, or the shell refused: not worth a dialog */ }
    }
}
