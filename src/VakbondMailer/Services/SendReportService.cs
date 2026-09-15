using System.Globalization;
using System.IO;
using System.Text;
using ClosedXML.Excel;
using CsvHelper;
using VakbondMailer.Models;

namespace VakbondMailer.Services;

public static class SendReportService
{
    private static readonly string[] Headers = ["Naam", "E-mail", "Status", "Foutmelding"];

    private static (string Naam, string Email, string Status, string Foutmelding) ToRow(SendResult result) =>
        (result.DisplayName, result.Email, result.Success ? "Verstuurd" : "Mislukt", result.Error ?? string.Empty);

    public static void Write(string filePath, IEnumerable<SendResult> results)
    {
        // UTF-8 mét BOM: zonder BOM opent Excel de CSV als ANSI en worden namen
        // met accenten (José, Renée) onleesbaar.
        using var writer = new StreamWriter(filePath, false, new UTF8Encoding(encoderShouldEmitUTF8Identifier: true));
        using var csv = new CsvWriter(writer, CultureInfo.InvariantCulture);

        foreach (var header in Headers)
            csv.WriteField(header);
        csv.NextRecord();

        foreach (var result in results)
        {
            var row = ToRow(result);
            csv.WriteField(row.Naam);
            csv.WriteField(row.Email);
            csv.WriteField(row.Status);
            csv.WriteField(row.Foutmelding);
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

        for (var i = 0; i < Headers.Length; i++)
            sheet.Cell(1, i + 1).Value = Headers[i];
        sheet.Range(1, 1, 1, Headers.Length).Style.Font.Bold = true;

        var rowIndex = 2;
        foreach (var result in results)
        {
            var row = ToRow(result);
            sheet.Cell(rowIndex, 1).Value = row.Naam;
            sheet.Cell(rowIndex, 2).Value = row.Email;

            var statusCell = sheet.Cell(rowIndex, 3);
            statusCell.Value = row.Status;
            statusCell.Style.Fill.BackgroundColor = result.Success
                ? XLColor.FromArgb(0xC6, 0xEF, 0xCE)
                : XLColor.FromArgb(0xFF, 0xC7, 0xCE);

            sheet.Cell(rowIndex, 4).Value = row.Foutmelding;
            rowIndex++;
        }

        sheet.Columns().AdjustToContents();
        workbook.SaveAs(filePath);
    }
}
