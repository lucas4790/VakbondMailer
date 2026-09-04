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
}
