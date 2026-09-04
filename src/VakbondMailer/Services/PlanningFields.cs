using System.Globalization;
using System.Linq;

namespace VakbondMailer.Services;

/// <summary>
/// Velden die niet uit de ledenlijst komen maar uit de app zelf: de maand waarin je een
/// gastles wilt inplannen, en eventueel een paar concrete datums om voor te stellen.
/// Ze zijn voor iedere ontvanger gelijk en zijn in het sjabloon te gebruiken als {{Maand}} enz.
/// </summary>
public static class PlanningFields
{
    public const string MonthKey = "Maand";
    public const string MonthYearKey = "MaandJaar";
    public const string DateOptionsKey = "Datumopties";

    private static readonly CultureInfo Dutch = new("nl-NL");

    public static IReadOnlyList<string> Keys { get; } = new[] { MonthKey, MonthYearKey, DateOptionsKey };

    public static string FormatMonth(DateTime month) => month.ToString("MMMM", Dutch);

    public static string FormatMonthYear(DateTime month) => month.ToString("MMMM yyyy", Dutch);

    /// <summary>
    /// Leest als een voorstel: "dinsdag 4 november, donderdag 13 november of maandag 24 november".
    /// </summary>
    public static string FormatDateOptions(IEnumerable<DateTime> dates)
    {
        var formatted = dates
            .Select(d => d.Date)
            .Distinct()
            .OrderBy(d => d)
            .Select(d => d.ToString("dddd d MMMM", Dutch))
            .ToList();

        return formatted.Count switch
        {
            0 => string.Empty,
            1 => formatted[0],
            _ => $"{string.Join(", ", formatted.Take(formatted.Count - 1))} of {formatted[^1]}",
        };
    }

    /// <summary>
    /// De eerste dag van de komende <paramref name="count"/> maanden, te beginnen bij de maand
    /// waar <paramref name="from"/> in valt — de keuzelijst om een gastles in te plannen.
    /// </summary>
    public static IReadOnlyList<DateTime> NextMonths(int count, DateTime from)
    {
        var firstOfMonth = new DateTime(from.Year, from.Month, 1);
        return Enumerable.Range(0, count).Select(firstOfMonth.AddMonths).ToList();
    }

    public static IReadOnlyDictionary<string, string> Build(DateTime month, IEnumerable<DateTime> dates) =>
        new Dictionary<string, string>
        {
            [MonthKey] = FormatMonth(month),
            [MonthYearKey] = FormatMonthYear(month),
            [DateOptionsKey] = FormatDateOptions(dates),
        };
}
