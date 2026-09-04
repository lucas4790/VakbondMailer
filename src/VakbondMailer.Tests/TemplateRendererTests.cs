using VakbondMailer.Models;
using VakbondMailer.Services;
using Xunit;

namespace VakbondMailer.Tests;

public class TemplateRendererTests
{
    private static Recipient CreateRecipient(params (string Key, string Value)[] fields)
    {
        var dict = fields.ToDictionary(f => f.Key, f => f.Value);
        return new Recipient
        {
            Email = dict.GetValueOrDefault("E-mail", "test@example.com"),
            Fields = dict,
        };
    }

    [Fact]
    public void Render_ReplacesKnownPlaceholder()
    {
        var recipient = CreateRecipient(("Voornaam", "Anne"));

        var result = TemplateRenderer.Render("Beste {{Voornaam}},", recipient);

        Assert.Equal("Beste Anne,", result);
    }

    [Fact]
    public void Render_LeavesUnknownPlaceholderUntouched()
    {
        var recipient = CreateRecipient(("Voornaam", "Anne"));

        var result = TemplateRenderer.Render("Beste {{Achternaam}},", recipient);

        Assert.Equal("Beste {{Achternaam}},", result);
    }

    [Fact]
    public void Render_TrimsWhitespaceInsidePlaceholder()
    {
        var recipient = CreateRecipient(("Voornaam", "Anne"));

        var result = TemplateRenderer.Render("Beste {{ Voornaam }},", recipient);

        Assert.Equal("Beste Anne,", result);
    }

    [Fact]
    public void Render_ReplacesMultiplePlaceholders()
    {
        var recipient = CreateRecipient(("Voornaam", "Anne"), ("Afdeling", "Zorg"));

        var result = TemplateRenderer.Render("Beste {{Voornaam}}, namens afdeling {{Afdeling}}.", recipient);

        Assert.Equal("Beste Anne, namens afdeling Zorg.", result);
    }

    [Fact]
    public void ExtractPlaceholders_ReturnsUniqueFieldNames()
    {
        var result = TemplateRenderer.ExtractPlaceholders("Beste {{Voornaam}}, {{Voornaam}} van {{School}}.");

        Assert.Equal(new[] { "Voornaam", "School" }, result);
    }

    [Fact]
    public void ExtractPlaceholders_ReturnsEmptyForPlainText()
    {
        Assert.Empty(TemplateRenderer.ExtractPlaceholders("Geen velden hier."));
    }

    [Fact]
    public void Render_TreatsDollarSignsInValuesAsPlainText()
    {
        // Regex-vervanging mag "$1" in een celwaarde niet als groepsverwijzing opvatten.
        var recipient = CreateRecipient(("Bedrag", "$1 per maand"));

        var result = TemplateRenderer.Render("Contributie: {{Bedrag}}", recipient);

        Assert.Equal("Contributie: $1 per maand", result);
    }

    [Fact]
    public void Render_DoesNotRecurseIntoValuesThatLookLikePlaceholders()
    {
        var recipient = CreateRecipient(("Voornaam", "{{Achternaam}}"), ("Achternaam", "de Boer"));

        var result = TemplateRenderer.Render("Beste {{Voornaam}}", recipient);

        Assert.Equal("Beste {{Achternaam}}", result);
    }
}
