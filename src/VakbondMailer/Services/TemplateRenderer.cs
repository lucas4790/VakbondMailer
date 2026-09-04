using System.Linq;
using System.Text.RegularExpressions;
using VakbondMailer.Models;

namespace VakbondMailer.Services;

public static partial class TemplateRenderer
{
    /// <param name="extraFields">
    /// Velden die niet uit de ledenlijst komen (zie <see cref="PlanningFields"/>). Kolommen uit
    /// de ledenlijst gaan voor: die zijn immers per ontvanger ingevuld.
    /// </param>
    public static string Render(string template, Recipient recipient, IReadOnlyDictionary<string, string>? extraFields = null)
    {
        return PlaceholderPattern().Replace(template, match =>
        {
            var key = match.Groups[1].Value.Trim();

            if (recipient.Fields.TryGetValue(key, out var value))
                return value;

            if (extraFields is not null && extraFields.TryGetValue(key, out var extra))
                return extra;

            return match.Value;
        });
    }

    /// <summary>
    /// Alle unieke veldnamen die als <c>{{Veld}}</c> in het sjabloon voorkomen, ongeacht of
    /// ze ook echt in de ledenlijst bestaan (dat wordt elders gecontroleerd).
    /// </summary>
    public static IReadOnlyList<string> ExtractPlaceholders(string template) =>
        PlaceholderPattern().Matches(template)
            .Select(m => m.Groups[1].Value.Trim())
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .ToList();

    /// <summary>
    /// Velden die in het sjabloon gebruikt worden maar nergens uit ingevuld kunnen worden — vaak
    /// een tikfout. Die blijven anders letterlijk als {{Voornam}} in de verstuurde mail staan.
    /// </summary>
    public static IReadOnlyList<string> FindUnknownPlaceholders(
        string subject,
        string body,
        IEnumerable<string> knownFields)
    {
        var known = knownFields.ToList();

        return ExtractPlaceholders(subject)
            .Concat(ExtractPlaceholders(body))
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .Where(token => !known.Contains(token, StringComparer.OrdinalIgnoreCase))
            .ToList();
    }

    [GeneratedRegex(@"\{\{\s*([^{}]+?)\s*\}\}")]
    private static partial Regex PlaceholderPattern();
}
