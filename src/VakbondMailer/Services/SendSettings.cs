using System.Globalization;

namespace VakbondMailer.Services;

/// <summary>
/// De regels rond het versturen: vanaf welk adres het mag, en hoe lang de app tussen twee
/// mails wacht.
/// </summary>
public static class SendSettings
{
    /// <summary>Er mag alleen verstuurd worden vanaf een adres op dit domein.</summary>
    public const string RequiredEmailDomain = "@fnv.nl";

    /// <summary>Binnen deze periode waarschuwt de app dat dezelfde mail al verstuurd is.</summary>
    public static readonly TimeSpan DuplicateSendWindow = TimeSpan.FromDays(14);

    private const double DefaultDelaySeconds = 1;

    /// <summary>Ondergrens: zonder pauze ziet Exchange een reeks mails al snel als spam.</summary>
    private const double MinimumDelaySeconds = 0.2;

    public static bool IsAllowedSender(string email) =>
        !string.IsNullOrWhiteSpace(email) && email.EndsWith(RequiredEmailDomain, StringComparison.OrdinalIgnoreCase);

    /// <summary>
    /// Leest de ingevulde pauze. Accepteert zowel "1,5" als "1.5" (op een Nederlandse Windows
    /// typt men een komma), en valt bij onleesbare invoer terug op de standaardwaarde.
    /// </summary>
    public static double ParseDelaySeconds(string text)
    {
        var normalized = (text ?? string.Empty).Trim().Replace(',', '.');

        if (double.TryParse(normalized, NumberStyles.Float, CultureInfo.InvariantCulture, out var seconds) && seconds >= 0)
            return Math.Max(seconds, MinimumDelaySeconds);

        return DefaultDelaySeconds;
    }
}
