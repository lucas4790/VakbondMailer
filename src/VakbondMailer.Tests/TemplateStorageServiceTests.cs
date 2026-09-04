using VakbondMailer.Services;
using Xunit;

namespace VakbondMailer.Tests;

public class TemplateStorageServiceTests
{
    [Fact]
    public void SuggestFileName_GebruiktHetOnderwerp()
    {
        Assert.Equal("Uitnodiging ledenvergadering", TemplateStorageService.SuggestFileName("Uitnodiging ledenvergadering"));
    }

    [Theory]
    [InlineData("")]
    [InlineData("   ")]
    public void SuggestFileName_ValtTerugOpEenStandaardnaam(string onderwerp)
    {
        Assert.Equal("standaardmail", TemplateStorageService.SuggestFileName(onderwerp));
    }

    [Fact]
    public void SuggestFileName_HaaltTekensWegDieWindowsNietToestaat()
    {
        var naam = TemplateStorageService.SuggestFileName("Cao: loon/uren?");

        Assert.DoesNotContain(":", naam);
        Assert.DoesNotContain("/", naam);
        Assert.DoesNotContain("?", naam);
        Assert.StartsWith("Cao", naam);
    }

    [Fact]
    public void SaveEnLoad_HoudenNaamOnderwerpEnTekstHeel()
    {
        var pad = Path.Combine(Path.GetTempPath(), $"{Guid.NewGuid()}.json");
        try
        {
            TemplateStorageService.Save(pad, new MailTemplate
            {
                Name = "Gastles inplannen",
                Subject = "Gastles in {{Maand}}",
                Body = "Beste {{Voornaam}},\n\nTot dan!",
            });

            var geladen = TemplateStorageService.Load(pad);

            Assert.Equal("Gastles inplannen", geladen.Name);
            Assert.Equal("Gastles in {{Maand}}", geladen.Subject);
            Assert.Contains("Tot dan!", geladen.Body);
        }
        finally
        {
            File.Delete(pad);
        }
    }
}
