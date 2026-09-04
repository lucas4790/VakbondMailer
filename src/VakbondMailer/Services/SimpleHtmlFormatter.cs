using System.Net;
using System.Text.RegularExpressions;

namespace VakbondMailer.Services;

/// <summary>
/// Zet platte tekst met eenvoudige markdown-achtige opmaak (**vet**, *cursief*) en kale
/// links om naar een minimale HTML-mail, voor wie liever opgemaakte mail verstuurt dan platte tekst.
/// </summary>
public static partial class SimpleHtmlFormatter
{
    public static string ToHtml(string plainText)
    {
        var escaped = WebUtility.HtmlEncode(plainText);
        var withBold = BoldPattern().Replace(escaped, "<b>$1</b>");
        var withItalic = ItalicPattern().Replace(withBold, "<i>$1</i>");
        var withLinks = UrlPattern().Replace(withItalic, m => $"<a href=\"{m.Value}\">{m.Value}</a>");
        var withBreaks = withLinks.Replace("\r\n", "\n").Replace("\n", "<br>\n");
        return $"<html><body style=\"font-family:Calibri,'Segoe UI',sans-serif;font-size:11pt\">{withBreaks}</body></html>";
    }

    [GeneratedRegex(@"\*\*(.+?)\*\*")]
    private static partial Regex BoldPattern();

    [GeneratedRegex(@"(?<!\*)\*(?!\*)(.+?)(?<!\*)\*(?!\*)")]
    private static partial Regex ItalicPattern();

    [GeneratedRegex(@"https?://[^\s<>""]+")]
    private static partial Regex UrlPattern();
}
