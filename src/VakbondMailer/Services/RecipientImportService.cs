using System.Globalization;
using System.IO;
using System.Linq;
using ClosedXML.Excel;
using CsvHelper;
using VakbondMailer.Models;

namespace VakbondMailer.Services;

public sealed class ImportedRecipients
{
    public required IReadOnlyList<string> Headers { get; init; }

    public required IReadOnlyList<Recipient> Recipients { get; init; }
}

public static class RecipientImportService
{
    public static IReadOnlyList<string> ReadHeaders(string filePath)
    {
        return GetExtension(filePath) switch
        {
            ".csv" => ReadCsvHeaders(filePath),
            ".xlsx" => ReadExcelHeaders(filePath),
            var ext => throw NotSupported(ext),
        };
    }

    public static ImportedRecipients Import(string filePath, string emailColumn)
    {
        return GetExtension(filePath) switch
        {
            ".csv" => ImportCsv(filePath, emailColumn),
            ".xlsx" => ImportExcel(filePath, emailColumn),
            var ext => throw NotSupported(ext),
        };
    }

    public static string? GuessEmailColumn(IReadOnlyList<string> headers)
    {
        return headers.FirstOrDefault(h =>
            h.Contains("e-mail", StringComparison.OrdinalIgnoreCase) ||
            h.Contains("email", StringComparison.OrdinalIgnoreCase) ||
            h.Equals("mail", StringComparison.OrdinalIgnoreCase));
    }

    private static string GetExtension(string filePath) => Path.GetExtension(filePath).ToLowerInvariant();

    private static NotSupportedException NotSupported(string extension) =>
        new($"Bestandstype '{extension}' wordt niet ondersteund. Gebruik een .csv- of .xlsx-bestand.");

    private static IReadOnlyList<string> ReadCsvHeaders(string filePath)
    {
        using var reader = new StreamReader(filePath);
        using var csv = new CsvReader(reader, CultureInfo.InvariantCulture);
        csv.Read();
        csv.ReadHeader();
        var headers = csv.HeaderRecord?.Select(h => h.Trim()).ToList() ?? new List<string>();
        EnsureUniqueHeaders(headers);
        return headers;
    }

    private static IReadOnlyList<string> ReadExcelHeaders(string filePath)
    {
        using var workbook = new XLWorkbook(filePath);
        return GetExcelColumnHeaders(workbook.Worksheets.First());
    }

    /// <summary>
    /// Leest headers als een 1:1 lijst per kolomnummer (met "KolomN" als fallback voor lege
    /// headercellen), zodat de positie in de lijst altijd overeenkomt met de kolomindex die
    /// verderop wordt gebruikt om celwaarden per rij op te halen.
    /// </summary>
    private static IReadOnlyList<string> GetExcelColumnHeaders(IXLWorksheet worksheet)
    {
        var headerRow = worksheet.FirstRowUsed()
            ?? throw new InvalidOperationException("Het Excel-bestand lijkt leeg te zijn.");
        var lastColumn = worksheet.LastColumnUsed()?.ColumnNumber() ?? 0;

        var headers = new List<string>(lastColumn);
        for (var col = 1; col <= lastColumn; col++)
        {
            var value = headerRow.Cell(col).GetString().Trim();
            headers.Add(value.Length > 0 ? value : $"Kolom{col}");
        }

        EnsureUniqueHeaders(headers);
        return headers;
    }

    private static void EnsureUniqueHeaders(IReadOnlyList<string> headers)
    {
        var duplicate = headers
            .GroupBy(h => h, StringComparer.OrdinalIgnoreCase)
            .FirstOrDefault(g => g.Count() > 1);
        if (duplicate is not null)
            throw new InvalidOperationException(
                $"De kolomnaam '{duplicate.Key}' komt meerdere keren voor. Zorg dat elke kolom een unieke naam heeft.");
    }

    private static ImportedRecipients ImportCsv(string filePath, string emailColumn)
    {
        using var reader = new StreamReader(filePath);
        using var csv = new CsvReader(reader, CultureInfo.InvariantCulture);
        csv.Read();
        csv.ReadHeader();
        var headers = csv.HeaderRecord?.Select(h => h.Trim()).ToList() ?? new List<string>();
        EnsureUniqueHeaders(headers);
        ValidateEmailColumn(headers, emailColumn);

        var recipients = new List<Recipient>();
        while (csv.Read())
        {
            var fields = headers.ToDictionary(h => h, h => csv.GetField(h)?.Trim() ?? string.Empty);
            var email = fields[emailColumn];
            if (string.IsNullOrWhiteSpace(email))
                continue;

            recipients.Add(new Recipient { Email = email, Fields = fields });
        }

        return new ImportedRecipients { Headers = headers, Recipients = recipients };
    }

    private static ImportedRecipients ImportExcel(string filePath, string emailColumn)
    {
        using var workbook = new XLWorkbook(filePath);
        var worksheet = workbook.Worksheets.First();
        var headers = GetExcelColumnHeaders(worksheet);
        ValidateEmailColumn(headers, emailColumn);

        var recipients = new List<Recipient>();
        foreach (var row in worksheet.RowsUsed().Skip(1))
        {
            var fields = new Dictionary<string, string>();
            for (var i = 0; i < headers.Count; i++)
            {
                fields[headers[i]] = row.Cell(i + 1).GetString().Trim();
            }

            var email = fields[emailColumn];
            if (string.IsNullOrWhiteSpace(email))
                continue;

            recipients.Add(new Recipient { Email = email, Fields = fields });
        }

        return new ImportedRecipients { Headers = headers, Recipients = recipients };
    }

    private static void ValidateEmailColumn(IReadOnlyList<string> headers, string emailColumn)
    {
        if (!headers.Contains(emailColumn))
            throw new ArgumentException($"Kolom '{emailColumn}' niet gevonden in het bestand.", nameof(emailColumn));
    }
}
