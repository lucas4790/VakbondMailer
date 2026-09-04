using System.Globalization;
using System.IO;
using CsvHelper;
using VakbondMailer.Models;

namespace VakbondMailer.Services;

public static class SendReportService
{
    public static void Write(string filePath, IEnumerable<SendResult> results)
    {
        using var writer = new StreamWriter(filePath);
        using var csv = new CsvWriter(writer, CultureInfo.InvariantCulture);

        csv.WriteField("Naam");
        csv.WriteField("E-mail");
        csv.WriteField("Status");
        csv.WriteField("Foutmelding");
        csv.NextRecord();

        foreach (var result in results)
        {
            csv.WriteField(result.DisplayName);
            csv.WriteField(result.Email);
            csv.WriteField(result.Success ? "Verstuurd" : "Mislukt");
            csv.WriteField(result.Error ?? string.Empty);
            csv.NextRecord();
        }
    }
}
