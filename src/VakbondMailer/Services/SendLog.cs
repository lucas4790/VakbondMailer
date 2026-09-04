using System.Collections.ObjectModel;

namespace VakbondMailer.Services;

/// <summary>
/// Eén regel in het logboek. De XAML-DataTemplate bindt op deze property-namen.
/// </summary>
public sealed record LogEntry(string Time, string Title, string? Detail, bool? Success)
{
    public bool HasDetail => !string.IsNullOrEmpty(Detail);

    public bool IsSuccess => Success == true;

    public bool IsFailure => Success == false;
}

/// <summary>
/// Het logboek dat tijdens het verzenden meeloopt: neutrale meldingen en per ontvanger of het
/// gelukt is.
/// </summary>
public sealed class SendLog
{
    private readonly Func<DateTime> _now;

    public SendLog(Func<DateTime>? now = null) => _now = now ?? (() => DateTime.Now);

    public ObservableCollection<LogEntry> Entries { get; } = new();

    public LogEntry Add(string message) => Append(new LogEntry(Stamp(), message, null, null));

    public LogEntry AddSend(string title, string email, bool success, string? error = null)
    {
        var detail = success ? email : $"{email} — {error}";
        return Append(new LogEntry(Stamp(), title, detail, success));
    }

    public void Clear() => Entries.Clear();

    private LogEntry Append(LogEntry entry)
    {
        Entries.Add(entry);
        return entry;
    }

    private string Stamp() => _now().ToString("HH:mm:ss");
}
