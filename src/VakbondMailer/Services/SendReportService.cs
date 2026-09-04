using System.Globalization;
using System.IO;
using System.Text;
using CsvHelper;
using VakbondMailer.Models;

namespace VakbondMailer.Services;

public static class SendReportService
{
    public static void Write(string filePath, IEnumerable<SendResult> results)
    {
        // UTF-8 mét BOM: zonder BOM opent Excel de CSV als ANSI en worden namen
        // met accenten (José, Renée) onleesbaar.
        using var writer = new StreamWriter(filePath, false, new UTF8Encoding(encoderShouldEmitUTF8Identifier: true));
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
