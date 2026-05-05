using System.Text.RegularExpressions;

namespace Outlook.ReminderApp;

internal static partial class TeamsJoinLinkResolver
{
    [GeneratedRegex(@"https://teams(?:\.live)?\.microsoft\.com/(?:l/meetup-join|meet)/[\w\-._~:/?#\[\]@!$&'()*+,;=%]+", RegexOptions.IgnoreCase)]
    private static partial Regex TeamsJoinRegex();

    [GeneratedRegex(@"https://teams(?:\.live)?\.microsoft\.com/l/(?:chat|message)/[\w\-._~:/?#\[\]@!$&'()*+,;=%]+", RegexOptions.IgnoreCase)]
    private static partial Regex TeamsChatRegex();

    // Extracts the thread ID segment from a meetup-join URL.
    // Captures everything between /l/meetup-join/ and the next / or end-of-relevant-path.
    [GeneratedRegex(@"/l/meetup-join/([^/\s""<>]+)/", RegexOptions.IgnoreCase)]
    private static partial Regex MeetupJoinThreadRegex();

    [GeneratedRegex(@"https://[a-z0-9]+\.safelinks\.protection\.outlook\.com/[^\s""<>]*?[?&]url=([^&\s""<>]+)", RegexOptions.IgnoreCase)]
    private static partial Regex SafeLinksRegex();

    [GeneratedRegex(@"[?&]context=([^&\s""<>]+)", RegexOptions.IgnoreCase)]
    private static partial Regex ContextParamRegex();

    [GeneratedRegex(@"""Tid""\s*:\s*""([^""]+)""", RegexOptions.IgnoreCase)]
    private static partial Regex TidRegex();

    public static string? Resolve(string? body, string? location)
    {
        var fromBody = TryExtract(body, TeamsJoinRegex());
        if (fromBody is not null)
        {
            return fromBody;
        }

        return TryExtract(location, TeamsJoinRegex());
    }

    public static string? ResolveChat(string? body, string? location)
    {
        var fromBody = TryExtract(body, TeamsChatRegex());
        if (fromBody is not null)
        {
            return fromBody;
        }

        return TryExtract(location, TeamsChatRegex());
    }

    /// <summary>
    /// Derives a Teams meeting chat deep-link from a join URL by extracting the thread ID.
    /// Falls back to null if the join URL doesn't contain a recognisable thread.
    /// </summary>
    public static string? DeriveChatUrlFromJoinUrl(string? joinUrl)
    {
        if (string.IsNullOrWhiteSpace(joinUrl)) return null;
        var match = MeetupJoinThreadRegex().Match(joinUrl);
        if (!match.Success) return null;
        var threadEncoded = match.Groups[1].Value;
        var chatUrl = $"https://teams.microsoft.com/l/chat/{threadEncoded}/";
        var tenantId = ExtractTenantId(joinUrl);
        return tenantId is not null ? $"{chatUrl}?tenantId={tenantId}" : chatUrl;
    }

    /// <summary>
    /// Tries URL-decoding the body once to surface meetup-join URLs wrapped inside safelinks,
    /// then derives a chat deep-link from the thread ID found there.
    /// </summary>
    public static string? DeriveChatUrlFromDecodedBody(string? body)
    {
        if (string.IsNullOrWhiteSpace(body)) return null;
        try
        {
            var decoded = Uri.UnescapeDataString(body);
            foreach (Match m in TeamsJoinRegex().Matches(decoded))
            {
                var chatUrl = DeriveChatUrlFromJoinUrl(m.Value);
                if (chatUrl is not null) return chatUrl;
            }
            return null;
        }
        catch
        {
            return null;
        }
    }

    private static string? ExtractTenantId(string? url)
    {
        if (string.IsNullOrWhiteSpace(url)) return null;
        var contextMatch = ContextParamRegex().Match(url);
        if (!contextMatch.Success) return null;
        try
        {
            var contextJson = Uri.UnescapeDataString(contextMatch.Groups[1].Value);
            // May be double-encoded (e.g. from SafeLinks extraction)
            if (contextJson.Contains('%'))
                contextJson = Uri.UnescapeDataString(contextJson);
            var tidMatch = TidRegex().Match(contextJson);
            return tidMatch.Success ? tidMatch.Groups[1].Value : null;
        }
        catch { return null; }
    }

    private static string? TryExtract(string? content, Regex regex)
    {
        if (string.IsNullOrWhiteSpace(content))
        {
            return null;
        }

        var match = regex.Match(content);
        if (match.Success) return match.Value;

        // Fall back: look inside SafeLinks-wrapped URLs
        var safeMatch = SafeLinksRegex().Match(content);
        while (safeMatch.Success)
        {
            try
            {
                var decoded = Uri.UnescapeDataString(safeMatch.Groups[1].Value);
                var inner = regex.Match(decoded);
                if (inner.Success) return inner.Value;
            }
            catch { }
            safeMatch = safeMatch.NextMatch();
        }

        return null;
    }
}