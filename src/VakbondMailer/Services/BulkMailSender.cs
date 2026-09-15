using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using VakbondMailer.Models;

namespace VakbondMailer.Services;

/// <summary>
/// Wat er per mail ingevuld moet worden. Wordt één keer vastgelegd vóór het verzenden, zodat
/// wijzigingen in het scherm de lopende verzending niet meer beïnvloeden.
/// </summary>
public sealed class BulkSendOptions
{
    public required string SubjectTemplate { get; init; }

    public required string BodyTemplate { get; init; }

    public IReadOnlyDictionary<string, string>? PlanningFields { get; init; }

    public bool IsHtml { get; init; }

    public IReadOnlyList<string>? AttachmentPaths { get; init; }

    public string? AccountName { get; init; }

    public TimeSpan DelayBetweenMails { get; init; } = TimeSpan.FromSeconds(1);
}

public sealed record BulkSendProgress(int Processed, int Total, Recipient Recipient, SendResult Result);

public sealed class BulkSendOutcome
{
    public required IReadOnlyList<SendResult> Results { get; init; }

    public required IReadOnlyList<Recipient> Failed { get; init; }

    public required IReadOnlyList<string> SentEmails { get; init; }

    public required bool Cancelled { get; init; }
}

/// <summary>
/// Verstuurt één mail per ontvanger, met een pauze ertussen, en houdt bij wat wel en niet
/// gelukt is. Losgetrokken van het scherm zodat dit stuk — waar het echt mis kan gaan —
/// getest kan worden.
/// </summary>
public static class BulkMailSender
{
    /// <remarks>
    /// Wordt bewust op de aanroepende (UI-)thread uitgevoerd: Outlook-COM-objecten zijn
    /// STA-gebonden. <paramref name="onProgress"/> wordt daarom ook direct aangeroepen.
    /// </remarks>
    public static async Task<BulkSendOutcome> SendAsync(
        IMailSender sender,
        IReadOnlyList<Recipient> recipients,
        BulkSendOptions options,
        Action<BulkSendProgress>? onProgress = null,
        CancellationToken cancellationToken = default)
    {
        var results = new List<SendResult>();
        var failed = new List<Recipient>();
        var sentEmails = new List<string>();
        var cancelled = false;

        for (var i = 0; i < recipients.Count; i++)
        {
            if (cancellationToken.IsCancellationRequested)
            {
                cancelled = true;
                break;
            }

            var recipient = recipients[i];
            var result = SendOne(sender, recipient, options);

            if (result.Success)
                sentEmails.Add(recipient.Email);
            else
                failed.Add(recipient);

            results.Add(result);
            onProgress?.Invoke(new BulkSendProgress(i + 1, recipients.Count, recipient, result));

            if (i < recipients.Count - 1)
            {
                try
                {
                    await Task.Delay(options.DelayBetweenMails, cancellationToken);
                }
                catch (TaskCanceledException)
                {
                    cancelled = true;
                    break;
                }
            }
        }

        return new BulkSendOutcome
        {
            Results = results,
            Failed = failed,
            SentEmails = sentEmails,
            Cancelled = cancelled,
        };
    }

    /// <summary>
    /// Eén mail. Een fout bij deze ontvanger mag de rest van de lijst niet tegenhouden, dus
    /// die wordt hier opgevangen en als resultaat teruggegeven.
    /// </summary>
    private static SendResult SendOne(IMailSender sender, Recipient recipient, BulkSendOptions options)
    {
        try
        {
            var subject = TemplateRenderer.Render(options.SubjectTemplate, recipient, options.PlanningFields);
            var body = ComposeBody(options, recipient);
            sender.SendMail(recipient.Email, subject, body, options.AccountName, options.IsHtml, options.AttachmentPaths);

            return new SendResult { Email = recipient.Email, DisplayName = recipient.DisplayName, Success = true };
        }
        catch (Exception ex)
        {
            return new SendResult
            {
                Email = recipient.Email,
                DisplayName = recipient.DisplayName,
                Success = false,
                Error = ex.Message,
            };
        }
    }

    public static string ComposeBody(BulkSendOptions options, Recipient recipient)
    {
        var rendered = TemplateRenderer.Render(options.BodyTemplate, recipient, options.PlanningFields);
        return options.IsHtml ? SimpleHtmlFormatter.ToHtml(rendered) : rendered;
    }
}
