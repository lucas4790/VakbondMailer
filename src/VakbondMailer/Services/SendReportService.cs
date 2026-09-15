using System.Globalization;
using System.IO;
using System.Text;
using ClosedXML.Excel;
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

    /// <summary>
    /// Zelfde inhoud als <see cref="Write"/>, maar als .xlsx met een gekleurde statuskolom —
    /// leesbaarder dan de .csv voor iemand die geen technische achtergrond heeft.
    /// </summary>
    public static void WriteXlsx(string filePath, IEnumerable<SendResult> results)
    {
        using var workbook = new XLWorkbook();
        var sheet = workbook.Worksheets.Add("Verzendrapport");

        sheet.Cell(1, 1).Value = "Naam";
        sheet.Cell(1, 2).Value = "E-mail";
        sheet.Cell(1, 3).Value = "Status";
        sheet.Cell(1, 4).Value = "Foutmelding";
        sheet.Range(1, 1, 1, 4).Style.Font.Bold = true;

        var row = 2;
        foreach (var result in results)
        {
            sheet.Cell(row, 1).Value = result.DisplayName;
            sheet.Cell(row, 2).Value = result.Email;

            var statusCell = sheet.Cell(row, 3);
            statusCell.Value = result.Success ? "Verstuurd" : "Mislukt";
            statusCell.Style.Fill.BackgroundColor = result.Success
                ? XLColor.FromArgb(0xC6, 0xEF, 0xCE)
                : XLColor.FromArgb(0xFF, 0xC7, 0xCE);

            sheet.Cell(row, 4).Value = result.Error ?? string.Empty;
            row++;
        }

        sheet.Columns().AdjustToContents();
        workbook.SaveAs(filePath);
    }
}
