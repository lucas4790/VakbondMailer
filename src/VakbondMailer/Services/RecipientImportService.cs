using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using ClosedXML.Excel;
using CsvHelper;
using VakbondMailer.Models;

namespace VakbondMailer.Services;

public sealed class ImportedRecipients
{
    public required IReadOnlyList<string> Headers { get; init; }

    public required IReadOnlyList<Recipient> Recipients { get; init; }

    /// <summary>
    /// Rijen die zijn overgeslagen (geen/ongeldig e-mailadres, of een dubbel adres), als
    /// leesbare tekst — zodat de gebruiker vóór het verzenden kan zien wat er is uitgesloten.
    /// </summary>
    public required IReadOnlyList<string> Warnings { get; init; }
}

public static partial class RecipientImportService
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

    /// <summary>
    /// Best-effort gok naar een naam- en schoolkolom, alleen gebruikt om dezelfde docent met
    /// twee verschillende e-mailadressen te kunnen signaleren. Ontbreekt een van beide, dan
    /// slaan we die extra check simpelweg over — de e-mail-duplicate-check blijft dan de enige.
    /// </summary>
    private static string? GuessNameColumn(IReadOnlyList<string> headers) => GuessColumn(headers, "naam", "name");

    private static string? GuessSchoolColumn(IReadOnlyList<string> headers) => GuessColumn(headers, "school");

    private static string? GuessColumn(IReadOnlyList<string> headers, params string[] keywords) =>
        headers.FirstOrDefault(h => keywords.Any(k => h.Contains(k, StringComparison.OrdinalIgnoreCase)));

    /// <summary>
    /// Opent een CSV met de juiste codering: met BOM leest StreamReader die zelf, zonder BOM
    /// proberen we strikt UTF-8 en vallen we terug op Latin1 — dat laatste is wat Excel schrijft
    /// bij "CSV (gescheiden door lijstscheidingstekens)", waar accenten anders onleesbaar worden.
    /// </summary>
    private static StreamReader OpenCsvReader(string filePath)
    {
        var bytes = File.ReadAllBytes(filePath);

        if (StartsWithBom(bytes))
            return new StreamReader(new MemoryStream(bytes), Encoding.UTF8, detectEncodingFromByteOrderMarks: true);

        try
        {
            new UTF8Encoding(false, throwOnInvalidBytes: true).GetString(bytes);
            return new StreamReader(new MemoryStream(bytes), new UTF8Encoding(false));
        }
        catch (DecoderFallbackException)
        {
            return new StreamReader(new MemoryStream(bytes), Encoding.Latin1);
        }
    }

    private static bool StartsWithBom(byte[] bytes) =>
        (bytes.Length >= 3 && bytes[0] == 0xEF && bytes[1] == 0xBB && bytes[2] == 0xBF)
        || (bytes.Length >= 2 && ((bytes[0] == 0xFF && bytes[1] == 0xFE) || (bytes[0] == 0xFE && bytes[1] == 0xFF)));

    private static string GetExtension(string filePath) => Path.GetExtension(filePath).ToLowerInvariant();

    private static NotSupportedException NotSupported(string extension) =>
        new($"Bestandstype '{extension}' wordt niet ondersteund. Gebruik een .csv- of .xlsx-bestand.");

    private static IReadOnlyList<string> ReadCsvHeaders(string filePath)
    {
        using var reader = OpenCsvReader(filePath);
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

    /// <summary>
    /// Alle mutable toestand die tijdens één import wordt opgebouwd, gebundeld zodat
    /// <see cref="ProcessRow"/> niet elk stukje daarvan als losse parameter hoeft door te geven.
    /// </summary>
    private sealed class ImportState
    {
        public required string EmailColumn { get; init; }

        public string? NameColumn { get; init; }

        public string? SchoolColumn { get; init; }

        public List<Recipient> Recipients { get; } = new();

        public List<string> Warnings { get; } = new();

        public HashSet<string> SeenEmails { get; } = new(StringComparer.OrdinalIgnoreCase);

        public Dictionary<string, string> SeenNameSchoolCombinations { get; } = new(StringComparer.OrdinalIgnoreCase);
    }

    private static ImportState CreateImportState(IReadOnlyList<string> headers, string emailColumn) => new()
    {
        EmailColumn = emailColumn,
        NameColumn = GuessNameColumn(headers),
        SchoolColumn = GuessSchoolColumn(headers),
    };

    private static ImportedRecipients ImportCsv(string filePath, string emailColumn)
    {
        using var reader = OpenCsvReader(filePath);
        using var csv = new CsvReader(reader, CultureInfo.InvariantCulture);
        csv.Read();
        csv.ReadHeader();
        var headers = csv.HeaderRecord?.Select(h => h.Trim()).ToList() ?? new List<string>();
        EnsureUniqueHeaders(headers);
        ValidateEmailColumn(headers, emailColumn);
        var state = CreateImportState(headers, emailColumn);

        var rowNumber = 1;
        while (csv.Read())
        {
            rowNumber++;
            var fields = headers.ToDictionary(h => h, h => csv.GetField(h)?.Trim() ?? string.Empty);
            ProcessRow(fields, rowNumber, state);
        }

        return new ImportedRecipients { Headers = headers, Recipients = state.Recipients, Warnings = state.Warnings };
    }

    private static ImportedRecipients ImportExcel(string filePath, string emailColumn)
    {
        using var workbook = new XLWorkbook(filePath);
        var worksheet = workbook.Worksheets.First();
        var headers = GetExcelColumnHeaders(worksheet);
        ValidateEmailColumn(headers, emailColumn);
        var state = CreateImportState(headers, emailColumn);

        foreach (var row in worksheet.RowsUsed().Skip(1))
        {
            var fields = new Dictionary<string, string>();
            for (var i = 0; i < headers.Count; i++)
            {
                fields[headers[i]] = row.Cell(i + 1).GetString().Trim();
            }

            ProcessRow(fields, row.RowNumber(), state);
        }

        return new ImportedRecipients { Headers = headers, Recipients = state.Recipients, Warnings = state.Warnings };
    }

    private static void ProcessRow(Dictionary<string, string> fields, int rowNumber, ImportState state)
    {
        var email = fields[state.EmailColumn];
        if (string.IsNullOrWhiteSpace(email))
        {
            state.Warnings.Add($"Rij {rowNumber}: geen e-mailadres, overgeslagen.");
            return;
        }

        if (!LooksLikeEmail(email))
        {
            state.Warnings.Add($"Rij {rowNumber}: '{email}' lijkt geen geldig e-mailadres, overgeslagen.");
            return;
        }

        if (!state.SeenEmails.Add(email))
        {
            state.Warnings.Add($"Rij {rowNumber}: '{email}' staat dubbel in de lijst, deze keer overgeslagen.");
            return;
        }

        WarnIfLikelySamePersonWithDifferentEmail(fields, email, rowNumber, state);

        state.Recipients.Add(new Recipient { Email = email, Fields = fields });
    }

    /// <summary>
    /// Vangt de situatie waarin dezelfde docent twee keer in de lijst staat met een ander
    /// e-mailadres (bv. een tikfout of een oud adres) — anders dan de exacte e-mail-duplicate
    /// hierboven wordt deze rij wél geïmporteerd, alleen met een waarschuwing erbij.
    /// </summary>
    private static void WarnIfLikelySamePersonWithDifferentEmail(
        Dictionary<string, string> fields, string email, int rowNumber, ImportState state)
    {
        if (state.NameColumn is null || state.SchoolColumn is null)
            return;

        var name = fields[state.NameColumn];
        var school = fields[state.SchoolColumn];
        if (string.IsNullOrWhiteSpace(name) || string.IsNullOrWhiteSpace(school))
            return;

        var key = $"{name}|{school}";
        if (state.SeenNameSchoolCombinations.TryGetValue(key, out var eerderGezienEmail))
        {
            if (!string.Equals(eerderGezienEmail, email, StringComparison.OrdinalIgnoreCase))
            {
                state.Warnings.Add(
                    $"Rij {rowNumber}: '{name}' bij '{school}' staat al eerder in de lijst met een ander e-mailadres — controleer of dit niet dezelfde persoon is.");
            }

            return;
        }

        state.SeenNameSchoolCombinations[key] = email;
    }

    private static bool LooksLikeEmail(string email) => EmailPattern().IsMatch(email);

    private static void ValidateEmailColumn(IReadOnlyList<string> headers, string emailColumn)
    {
        if (!headers.Contains(emailColumn))
            throw new ArgumentException($"Kolom '{emailColumn}' niet gevonden in het bestand.", nameof(emailColumn));
    }

    // Komma's, puntkomma's en punthaken worden bewust geweigerd: Outlook ziet die als scheidingsteken
    // tussen ontvangers, dus één slordige cel mag nooit stilletijd naar meerdere mensen mailen.
    [GeneratedRegex(@"^[^@\s,;<>""]+@[^@\s,;<>""]+\.[^@\s,;<>""]+$")]
    private static partial Regex EmailPattern();
}
