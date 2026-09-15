using System.Text;
using ClosedXML.Excel;
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

    /// <summary>Bouwt een tijdelijk .xlsx-bestand: elke binnenste array is één rij, eerste rij is de kop.</summary>
    private static string CreateTempXlsx(params string[][] rows)
    {
        var path = Path.Combine(Path.GetTempPath(), $"{Guid.NewGuid()}.xlsx");
        using var workbook = new XLWorkbook();
        var worksheet = workbook.Worksheets.Add("Ledenlijst");

        for (var rowIndex = 0; rowIndex < rows.Length; rowIndex++)
        {
            var row = rows[rowIndex];
            for (var colIndex = 0; colIndex < row.Length; colIndex++)
                worksheet.Cell(rowIndex + 1, colIndex + 1).Value = row[colIndex];
        }

        workbook.SaveAs(path);
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
    public void Import_WarnsWhenSameNameAndSchoolHaveDifferentEmail()
    {
        var path = CreateTempCsv(
            "Naam,School,E-mail\n" +
            "Anne Jansen,De Vlinder,anne@school-a.nl\n" +
            "Anne Jansen,De Vlinder,anne.jansen@school-a.nl\n");

        try
        {
            var imported = RecipientImportService.Import(path, "E-mail");

            Assert.Equal(2, imported.Recipients.Count);
            Assert.Single(imported.Warnings);
            Assert.Contains("Anne Jansen", imported.Warnings[0]);
            Assert.Contains("De Vlinder", imported.Warnings[0]);
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void Import_DoesNotWarnWithoutNameOrSchoolColumn()
    {
        var path = CreateTempCsv(
            "Voornaam,E-mail\n" +
            "Anne,anne@school-a.nl\n" +
            "Bram,bram@school-a.nl\n");

        try
        {
            var imported = RecipientImportService.Import(path, "E-mail");

            Assert.Equal(2, imported.Recipients.Count);
            Assert.Empty(imported.Warnings);
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void Import_DoesNotWarnWhenSameNameButDifferentSchool()
    {
        var path = CreateTempCsv(
            "Naam,School,E-mail\n" +
            "Anne Jansen,De Vlinder,anne@school-a.nl\n" +
            "Anne Jansen,De Zonnebloem,anne@school-b.nl\n");

        try
        {
            var imported = RecipientImportService.Import(path, "E-mail");

            Assert.Equal(2, imported.Recipients.Count);
            Assert.Empty(imported.Warnings);
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

    [Fact]
    public void Import_Xlsx_ParsesRecipientsAndSkipsEmptyEmail()
    {
        var path = CreateTempXlsx(
            new[] { "Voornaam", "E-mail" },
            new[] { "Anne", "anne@example.com" },
            new[] { "Bram", "" },
            new[] { "Carla", "carla@example.com" });

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
    public void Import_Xlsx_SkipsInvalidEmailAndReportsWarning()
    {
        var path = CreateTempXlsx(
            new[] { "Voornaam", "E-mail" },
            new[] { "Anne", "anne@example.com" },
            new[] { "Bram", "niet-een-adres" });

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
    public void Import_Xlsx_SkipsDuplicateEmailAndReportsWarning()
    {
        var path = CreateTempXlsx(
            new[] { "Voornaam", "E-mail" },
            new[] { "Anne", "anne@example.com" },
            new[] { "Anne (nogmaals)", "ANNE@example.com" });

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
    public void Import_Xlsx_ThrowsWhenEmailColumnMissing()
    {
        var path = CreateTempXlsx(
            new[] { "Voornaam", "Telefoon" },
            new[] { "Anne", "0612345678" });

        try
        {
            Assert.Throws<ArgumentException>(() => RecipientImportService.Import(path, "E-mail"));
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void Import_Xlsx_EmptyHeaderCellGetsColumnFallbackName()
    {
        var path = CreateTempXlsx(
            new[] { "Voornaam", "", "E-mail" },
            new[] { "Anne", "iets", "anne@example.com" });

        try
        {
            var headers = RecipientImportService.ReadHeaders(path);

            Assert.Equal(new[] { "Voornaam", "Kolom2", "E-mail" }, headers);
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void Import_Xlsx_ThrowsWhenColumnNameIsDuplicated()
    {
        var path = CreateTempXlsx(
            new[] { "E-mail", "E-mail" },
            new[] { "anne@example.com", "anne@example.com" });

        try
        {
            Assert.Throws<InvalidOperationException>(() => RecipientImportService.ReadHeaders(path));
        }
        finally
        {
            File.Delete(path);
        }
    }
}
