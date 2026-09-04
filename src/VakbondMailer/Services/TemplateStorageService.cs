using System.IO;
using System.Text.Json;

namespace VakbondMailer.Services;

public sealed class MailTemplate
{
    public string Name { get; set; } = string.Empty;

    public string Subject { get; set; } = string.Empty;

    public string Body { get; set; } = string.Empty;
}

public static class TemplateStorageService
{
    private static readonly JsonSerializerOptions Options = new() { WriteIndented = true };

    public static void Save(string filePath, MailTemplate template)
    {
        File.WriteAllText(filePath, JsonSerializer.Serialize(template, Options));
    }

    public static MailTemplate Load(string filePath)
    {
        var json = File.ReadAllText(filePath);
        return JsonSerializer.Deserialize<MailTemplate>(json)
            ?? throw new InvalidOperationException("Kon het sjabloon-bestand niet lezen.");
    }

    /// <summary>
    /// Een bestandsnaam-voorstel op basis van het onderwerp, met tekens die Windows niet in een
    /// bestandsnaam toestaat vervangen door een streepje.
    /// </summary>
    public static string SuggestFileName(string subject)
    {
        var name = string.IsNullOrWhiteSpace(subject) ? "standaardmail" : subject;
        foreach (var invalid in Path.GetInvalidFileNameChars())
            name = name.Replace(invalid, '-');

        return name;
    }
}
