using System.IO;
using System.Text.Json;

namespace VakbondMailer.Services;

public sealed class AppSettings
{
    public string? TemplatesFolder { get; set; }
}

/// <summary>
/// Onthoudt kleine, niet-gevoelige app-instellingen (zoals de laatst gekozen sjablonenmap)
/// tussen sessies, in de persoonlijke AppData-map van de gebruiker.
/// </summary>
public static class AppSettingsService
{
    private static readonly JsonSerializerOptions Options = new() { WriteIndented = true };

    private static readonly string SettingsPath = Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
        "VakbondMailer", "settings.json");

    public static AppSettings Load()
    {
        try
        {
            if (!File.Exists(SettingsPath))
                return new AppSettings();

            var json = File.ReadAllText(SettingsPath);
            return JsonSerializer.Deserialize<AppSettings>(json) ?? new AppSettings();
        }
        catch
        {
            return new AppSettings();
        }
    }

    public static void Save(AppSettings settings)
    {
        var directory = Path.GetDirectoryName(SettingsPath);
        if (!string.IsNullOrEmpty(directory))
            Directory.CreateDirectory(directory);

        File.WriteAllText(SettingsPath, JsonSerializer.Serialize(settings, Options));
    }
}
