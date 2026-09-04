using System.Linq;
using System.Text.RegularExpressions;
using VakbondMailer.Models;

namespace VakbondMailer.Services;

public static partial class TemplateRenderer
{
    public static string Render(string template, Recipient recipient)
    {
        return PlaceholderPattern().Replace(template, match =>
        {
            var key = match.Groups[1].Value.Trim();
            return recipient.Fields.TryGetValue(key, out var value) ? value : match.Value;
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

    [GeneratedRegex(@"\{\{\s*([^{}]+?)\s*\}\}")]
    private static partial Regex PlaceholderPattern();
}
