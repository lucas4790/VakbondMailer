using System.Text;
using ClosedXML.Excel;
using VakbondMailer.Models;
using VakbondMailer.Services;
using Xunit;

namespace VakbondMailer.Tests;

public class SendReportServiceTests
{
    private static string TempCsvPath() => Path.Combine(Path.GetTempPath(), $"{Guid.NewGuid()}.csv");

    private static string TempXlsxPath() => Path.Combine(Path.GetTempPath(), $"{Guid.NewGuid()}.xlsx");

    [Fact]
    public void Write_StartsWithUtf8Bom_SoExcelShowsAccentsCorrectly()
    {
        var path = TempCsvPath();
        try
        {
            SendReportService.Write(path, new[]
            {
                new SendResult { Email = "renee@voorbeeld.nl", DisplayName = "Renée Müller", Success = true },
            });

            var bytes = File.ReadAllBytes(path);
            Assert.True(bytes.Length >= 3);
            Assert.Equal(0xEF, bytes[0]);
            Assert.Equal(0xBB, bytes[1]);
            Assert.Equal(0xBF, bytes[2]);
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void Write_RoundTripsAccentedNames()
    {
        var path = TempCsvPath();
        try
        {
            SendReportService.Write(path, new[]
            {
                new SendResult { Email = "renee@voorbeeld.nl", DisplayName = "Renée Müller", Success = true },
            });

            var text = File.ReadAllText(path, Encoding.UTF8);
            Assert.Contains("Renée Müller", text);
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void Write_RecordsFailureWithErrorMessage()
    {
        var path = TempCsvPath();
        try
        {
            SendReportService.Write(path, new[]
            {
                new SendResult { Email = "a@voorbeeld.nl", DisplayName = "Anne", Success = true },
                new SendResult { Email = "b@voorbeeld.nl", DisplayName = "Bram", Success = false, Error = "Postvak vol" },
            });

            var lines = File.ReadAllLines(path, Encoding.UTF8);
            Assert.Equal(3, lines.Length); // kop + 2 rijen
            Assert.Contains("Verstuurd", lines[1]);
            Assert.Contains("Mislukt", lines[2]);
            Assert.Contains("Postvak vol", lines[2]);
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void WriteXlsx_WritesHeaderAndRowsWithCorrectValues()
    {
        var path = TempXlsxPath();
        try
        {
            SendReportService.WriteXlsx(path, new[]
            {
                new SendResult { Email = "a@voorbeeld.nl", DisplayName = "Anne", Success = true },
                new SendResult { Email = "b@voorbeeld.nl", DisplayName = "Bram", Success = false, Error = "Postvak vol" },
            });

            using var workbook = new XLWorkbook(path);
            var sheet = workbook.Worksheets.First();

            Assert.Equal("Naam", sheet.Cell(1, 1).GetString());
            Assert.Equal("E-mail", sheet.Cell(1, 2).GetString());
            Assert.Equal("Status", sheet.Cell(1, 3).GetString());
            Assert.Equal("Foutmelding", sheet.Cell(1, 4).GetString());

            Assert.Equal("Anne", sheet.Cell(2, 1).GetString());
            Assert.Equal("a@voorbeeld.nl", sheet.Cell(2, 2).GetString());
            Assert.Equal("Verstuurd", sheet.Cell(2, 3).GetString());
            Assert.Equal(string.Empty, sheet.Cell(2, 4).GetString());

            Assert.Equal("Bram", sheet.Cell(3, 1).GetString());
            Assert.Equal("Mislukt", sheet.Cell(3, 3).GetString());
            Assert.Equal("Postvak vol", sheet.Cell(3, 4).GetString());
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void WriteXlsx_ColorsStatusCellBySuccess()
    {
        var path = TempXlsxPath();
        try
        {
            SendReportService.WriteXlsx(path, new[]
            {
                new SendResult { Email = "a@voorbeeld.nl", DisplayName = "Anne", Success = true },
                new SendResult { Email = "b@voorbeeld.nl", DisplayName = "Bram", Success = false, Error = "Postvak vol" },
            });

            using var workbook = new XLWorkbook(path);
            var sheet = workbook.Worksheets.First();

            var successColor = sheet.Cell(2, 3).Style.Fill.BackgroundColor;
            var failureColor = sheet.Cell(3, 3).Style.Fill.BackgroundColor;

            Assert.NotEqual(successColor, failureColor);
        }
        finally
        {
            File.Delete(path);
        }
    }
}
