using VakbondMailer.Services;
using Xunit;

namespace VakbondMailer.Tests;

public class RecipientImportServiceTests
{
    private static string CreateTempCsv(string content)
    {
        var path = Path.Combine(Path.GetTempPath(), $"{Guid.NewGuid()}.csv");
        File.WriteAllText(path, content);
        return path;
    }

    [Fact]
    public void Import_Csv_ParsesRecipientsAndSkipsEmptyEmail()
    {
        var path = CreateTempCsv(
            "Voornaam,E-mail\n" +
            "Anne,anne@example.com\n" +
            "Bram,\n" +
            "Carla,carla@example.com\n");

        try
        {
            var imported = RecipientImportService.Import(path, "E-mail");

            Assert.Equal(new[] { "Voornaam", "E-mail" }, imported.Headers);
            Assert.Equal(2, imported.Recipients.Count);
            Assert.Equal("anne@example.com", imported.Recipients[0].Email);
            Assert.Equal("carla@example.com", imported.Recipients[1].Email);
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void Import_Csv_ThrowsWhenEmailColumnMissing()
    {
        var path = CreateTempCsv("Voornaam,Telefoon\nAnne,0612345678\n");

        try
        {
            Assert.Throws<ArgumentException>(() => RecipientImportService.Import(path, "E-mail"));
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Theory]
    [InlineData(new[] { "Voornaam", "E-mail" }, "E-mail")]
    [InlineData(new[] { "Naam", "Email adres" }, "Email adres")]
    [InlineData(new[] { "Naam", "Mail" }, "Mail")]
    public void GuessEmailColumn_FindsLikelyColumn(string[] headers, string expected)
    {
        Assert.Equal(expected, RecipientImportService.GuessEmailColumn(headers));
    }

    [Fact]
    public void GuessEmailColumn_ReturnsNullWhenNoMatch()
    {
        Assert.Null(RecipientImportService.GuessEmailColumn(new[] { "Voornaam", "Telefoon" }));
    }
}
