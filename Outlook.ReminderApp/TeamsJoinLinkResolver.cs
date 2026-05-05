using System.Text.RegularExpressions;

namespace Outlook.ReminderApp;

internal static partial class TeamsJoinLinkResolver
{
    [GeneratedRegex(@"https://teams(?:\.live)?\.microsoft\.com/(?:l/meetup-join|meet)/[\w\-._~:/?#\[\]@!$&'()*+,;=%]+", RegexOptions.IgnoreCase)]
    private static partial Regex TeamsJoinRegex();

    public static string? Resolve(string? body, string? location)
    {
        var fromBody = TryExtract(body);
        if (fromBody is not null)
        {
            return fromBody;
        }

        return TryExtract(location);
    }

    private static string? TryExtract(string? content)
    {
        if (string.IsNullOrWhiteSpace(content))
        {
            return null;
        }

        var match = TeamsJoinRegex().Match(content);
        return match.Success ? match.Value : null;
    }
}