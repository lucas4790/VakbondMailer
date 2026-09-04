using System.Text;
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

    private static string CreateTempCsv(string content, Encoding encoding)
    {
        var path = Path.Combine(Path.GetTempPath(), $"{Guid.NewGuid()}.csv");
        File.WriteAllText(path, content, encoding);
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
            Assert.Single(imported.Warnings);
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void Import_Csv_SkipsInvalidEmailAndReportsWarning()
    {
        var path = CreateTempCsv(
            "Voornaam,E-mail\n" +
            "Anne,anne@example.com\n" +
            "Bram,niet-een-adres\n");

        try
        {
            var imported = RecipientImportService.Import(path, "E-mail");

            Assert.Single(imported.Recipients);
            Assert.Equal("anne@example.com", imported.Recipients[0].Email);
            Assert.Single(imported.Warnings);
            Assert.Contains("niet-een-adres", imported.Warnings[0]);
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void Import_Csv_SkipsDuplicateEmailAndReportsWarning()
    {
        var path = CreateTempCsv(
            "Voornaam,E-mail\n" +
            "Anne,anne@example.com\n" +
            "Anne (nogmaals),ANNE@example.com\n");

        try
        {
            var imported = RecipientImportService.Import(path, "E-mail");

            Assert.Single(imported.Recipients);
            Assert.Equal("anne@example.com", imported.Recipients[0].Email);
            Assert.Single(imported.Warnings);
            Assert.Contains("dubbel", imported.Warnings[0]);
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

    [Fact]
    public void Import_ReadsAccentsFromAnsiCsvAsExcelWritesThem()
    {
        // Excel schrijft bij "CSV (gescheiden door lijstscheidingstekens)" geen UTF-8 maar ANSI.
        var path = CreateTempCsv(
            "Voornaam,E-mail\n" +
            "Renée,renee@voorbeeld.nl\n",
            Encoding.Latin1);

        try
        {
            var imported = RecipientImportService.Import(path, "E-mail");

            Assert.Single(imported.Recipients);
            Assert.Equal("Renée", imported.Recipients[0].Fields["Voornaam"]);
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void Import_ReadsAccentsFromUtf8CsvWithBom()
    {
        var path = CreateTempCsv(
            "Voornaam,E-mail\n" +
            "Renée,renee@voorbeeld.nl\n",
            new UTF8Encoding(encoderShouldEmitUTF8Identifier: true));

        try
        {
            var imported = RecipientImportService.Import(path, "E-mail");

            Assert.Single(imported.Recipients);
            Assert.Equal("Renée", imported.Recipients[0].Fields["Voornaam"]);
            Assert.Equal("Voornaam", imported.Headers[0]); // BOM hoort niet in de kolomnaam te blijven staan
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Theory]
    [InlineData("a@x.nl;b@y.nl")]
    [InlineData("\"a@x.nl,b@y.nl\"")]   // komma in één cel hoort in CSV tussen aanhalingstekens
    [InlineData("\"Anne <anne@x.nl>\"")]
    public void Import_RejectsAddressesThatCouldReachMultiplePeople(string address)
    {
        var path = CreateTempCsv("Voornaam,E-mail\nAnne," + address + "\n");

        try
        {
            var imported = RecipientImportService.Import(path, "E-mail");

            Assert.Empty(imported.Recipients);
            Assert.Single(imported.Warnings);
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void Import_WarningsPointAtTheRowNumberFromTheFile()
    {
        var path = CreateTempCsv(
            "Voornaam,E-mail\n" +
            "Anne,anne@voorbeeld.nl\n" +
            "Bram,\n");

        try
        {
            var imported = RecipientImportService.Import(path, "E-mail");

            // Rij 1 is de kop, Anne is rij 2, Bram is rij 3.
            Assert.Contains("Rij 3", imported.Warnings[0]);
        }
        finally
        {
            File.Delete(path);
        }
    }
}
