using System.Text;
using VakbondMailer.Models;
using VakbondMailer.Services;
using Xunit;

namespace VakbondMailer.Tests;

public class SendReportServiceTests
{
    private static string TempCsvPath() => Path.Combine(Path.GetTempPath(), $"{Guid.NewGuid()}.csv");

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
}
