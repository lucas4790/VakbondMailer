using System.IO;
using System.Linq;

namespace VakbondMailer.Services;

public sealed class TemplateSummary
{
    public required string FilePath { get; init; }

    public required string DisplayName { get; init; }
}

/// <summary>
/// Leest alle sjablonen (.json) uit een door de gebruiker gekozen map, zodat standaardmails
/// over verschillende onderwerpen naast elkaar bewaard en snel gekozen kunnen worden.
/// </summary>
public static class TemplateLibraryService
{
    public static IReadOnlyList<TemplateSummary> ListTemplates(string folderPath)
    {
        if (!Directory.Exists(folderPath))
            return Array.Empty<TemplateSummary>();

        var summaries = new List<TemplateSummary>();
        foreach (var file in Directory.EnumerateFiles(folderPath, "*.json"))
        {
            string displayName;
            try
            {
                var template = TemplateStorageService.Load(file);
                displayName = string.IsNullOrWhiteSpace(template.Name)
                    ? Path.GetFileNameWithoutExtension(file)
                    : template.Name;
            }
            catch
            {
                displayName = Path.GetFileNameWithoutExtension(file);
            }

            summaries.Add(new TemplateSummary { FilePath = file, DisplayName = displayName });
        }

        return summaries.OrderBy(s => s.DisplayName, StringComparer.CurrentCultureIgnoreCase).ToList();
    }
}
