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

    [GeneratedRegex(@"\{\{\s*([^{}]+?)\s*\}\}")]
    private static partial Regex PlaceholderPattern();
}
