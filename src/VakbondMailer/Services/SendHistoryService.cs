using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Text.Json;

namespace VakbondMailer.Services;

public sealed record SendHistoryEntry(DateTime SentAt, string Subject, string RecipientHash);

/// <summary>
/// Houdt bij welke mail wanneer naar wie ging, zodat de app kan waarschuwen als dezelfde
/// mailing kort geleden al naar (een deel van) dezelfde mensen is gegaan.
///
/// Adressen worden bewust als hash bewaard en niet leesbaar: genoeg om te herkennen dat
/// iemand die mail al kreeg, zonder een ledenlijst op schijf achter te laten.
/// </summary>
public static class SendHistoryService
{
    private static readonly TimeSpan KeepHistoryFor = TimeSpan.FromDays(90);

    public static string DefaultPath { get; } = Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
        "VakbondMailer", "verzendgeschiedenis.json");

    public static IReadOnlyList<SendHistoryEntry> Load(string path)
    {
        try
        {
            if (!File.Exists(path))
                return Array.Empty<SendHistoryEntry>();

            return JsonSerializer.Deserialize<List<SendHistoryEntry>>(File.ReadAllText(path))
                ?? new List<SendHistoryEntry>();
        }
        catch
        {
            // Een onleesbare geschiedenis mag het verzenden nooit blokkeren.
            return Array.Empty<SendHistoryEntry>();
        }
    }

    public static void Append(string path, string subject, IEnumerable<string> emails, DateTime sentAt)
    {
        var entries = Load(path).ToList();
        entries.AddRange(emails.Select(email => new SendHistoryEntry(sentAt, subject, HashEmail(email))));

        var cutoff = sentAt - KeepHistoryFor;
        var pruned = entries.Where(entry => entry.SentAt >= cutoff).ToList();

        var directory = Path.GetDirectoryName(path);
        if (!string.IsNullOrEmpty(directory))
            Directory.CreateDirectory(directory);

        File.WriteAllText(path, JsonSerializer.Serialize(pruned, new JsonSerializerOptions { WriteIndented = true }));
    }

    /// <summary>
    /// Hoeveel van deze ontvangers dezelfde mailing al binnen <paramref name="window"/> kregen.
    /// </summary>
    public static int CountRecentlySent(string path, string subject, IEnumerable<string> emails, TimeSpan window, DateTime now)
    {
        var cutoff = now - window;
        var alreadySent = Load(path)
            .Where(entry => entry.SentAt >= cutoff && string.Equals(entry.Subject, subject, StringComparison.Ordinal))
            .Select(entry => entry.RecipientHash)
            .ToHashSet(StringComparer.Ordinal);

        if (alreadySent.Count == 0)
            return 0;

        return emails.Select(HashEmail).Distinct(StringComparer.Ordinal).Count(alreadySent.Contains);
    }

    private static string HashEmail(string email) =>
        Convert.ToHexString(SHA256.HashData(Encoding.UTF8.GetBytes(email.Trim().ToLowerInvariant())));
}
