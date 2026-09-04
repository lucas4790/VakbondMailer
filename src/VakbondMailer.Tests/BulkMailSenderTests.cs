using VakbondMailer.Models;
using VakbondMailer.Services;
using Xunit;

namespace VakbondMailer.Tests;

public class BulkMailSenderTests
{
    /// <summary>Onthoudt wat er verstuurd zou zijn, en kan op commando falen.</summary>
    private sealed class FakeMailSender : IMailSender
    {
        private readonly Func<string, bool>? _failFor;

        public FakeMailSender(Func<string, bool>? failFor = null) => _failFor = failFor;

        public List<(string To, string Subject, string Body)> Verstuurd { get; } = new();

        public void SendMail(string toEmail, string subject, string body, string? accountName = null,
            bool isHtml = false, IReadOnlyList<string>? attachmentPaths = null)
        {
            if (_failFor?.Invoke(toEmail) == true)
                throw new InvalidOperationException("Postvak vol");

            Verstuurd.Add((toEmail, subject, body));
        }
    }

    private static Recipient Ontvanger(string voornaam, string email) => new()
    {
        Email = email,
        Fields = new Dictionary<string, string> { ["Voornaam"] = voornaam },
    };

    private static BulkSendOptions Opties(string subject = "Gastles", string body = "Beste {{Voornaam}}") => new()
    {
        SubjectTemplate = subject,
        BodyTemplate = body,
        DelayBetweenMails = TimeSpan.Zero,
    };

    [Fact]
    public async Task SendAsync_MailtIedereenEnVultDeVeldenIn()
    {
        var sender = new FakeMailSender();
        var ontvangers = new[] { Ontvanger("Anne", "anne@school.nl"), Ontvanger("Bram", "bram@school.nl") };

        var outcome = await BulkMailSender.SendAsync(sender, ontvangers, Opties(),
            cancellationToken: TestContext.Current.CancellationToken);

        Assert.Equal(2, sender.Verstuurd.Count);
        Assert.Equal("Beste Anne", sender.Verstuurd[0].Body);
        Assert.Equal("Beste Bram", sender.Verstuurd[1].Body);
        Assert.All(outcome.Results, r => Assert.True(r.Success));
        Assert.Empty(outcome.Failed);
        Assert.False(outcome.Cancelled);
    }

    [Fact]
    public async Task SendAsync_HoudtMisluktenApartMaarGaatDoorMetDeRest()
    {
        var sender = new FakeMailSender(failFor: email => email == "bram@school.nl");
        var ontvangers = new[]
        {
            Ontvanger("Anne", "anne@school.nl"),
            Ontvanger("Bram", "bram@school.nl"),
            Ontvanger("Carla", "carla@school.nl"),
        };

        var outcome = await BulkMailSender.SendAsync(sender, ontvangers, Opties(),
            cancellationToken: TestContext.Current.CancellationToken);

        Assert.Equal(3, outcome.Results.Count);
        Assert.Equal(2, sender.Verstuurd.Count); // Bram is niet aangekomen
        Assert.Equal(new[] { "anne@school.nl", "carla@school.nl" }, outcome.SentEmails);

        var mislukt = Assert.Single(outcome.Failed);
        Assert.Equal("bram@school.nl", mislukt.Email);
        Assert.Equal("Postvak vol", outcome.Results[1].Error);
    }

    [Fact]
    public async Task SendAsync_StoptDirectBijAnnuleren()
    {
        var sender = new FakeMailSender();
        var cts = new CancellationTokenSource();
        var ontvangers = Enumerable.Range(1, 5)
            .Select(i => Ontvanger($"Docent{i}", $"docent{i}@school.nl"))
            .ToList();

        // Annuleer zodra de eerste mail eruit is.
        var outcome = await BulkMailSender.SendAsync(sender, ontvangers, Opties(),
            onProgress: _ => cts.Cancel(), cancellationToken: cts.Token);

        Assert.True(outcome.Cancelled);
        Assert.Single(sender.Verstuurd);
        Assert.Single(outcome.Results);
    }

    [Fact]
    public async Task SendAsync_VerstuurtNietsWanneerVoorafAlGeannuleerdIs()
    {
        var sender = new FakeMailSender();
        var cts = new CancellationTokenSource();
        cts.Cancel();

        var outcome = await BulkMailSender.SendAsync(sender, new[] { Ontvanger("Anne", "anne@school.nl") },
            Opties(), cancellationToken: cts.Token);

        Assert.Empty(sender.Verstuurd);
        Assert.True(outcome.Cancelled);
    }

    [Fact]
    public async Task SendAsync_MeldtVoortgangPerOntvangerInVolgorde()
    {
        var sender = new FakeMailSender();
        var ontvangers = new[] { Ontvanger("Anne", "anne@school.nl"), Ontvanger("Bram", "bram@school.nl") };
        var voortgang = new List<string>();

        await BulkMailSender.SendAsync(sender, ontvangers, Opties(),
            onProgress: p => voortgang.Add($"{p.Processed}/{p.Total} {p.Recipient.Email}"),
            cancellationToken: TestContext.Current.CancellationToken);

        Assert.Equal(new[] { "1/2 anne@school.nl", "2/2 bram@school.nl" }, voortgang);
    }

    [Fact]
    public async Task SendAsync_GebruiktPlanningsveldenNaastDeLijstkolommen()
    {
        var sender = new FakeMailSender();
        var opties = new BulkSendOptions
        {
            SubjectTemplate = "Gastles in {{Maand}}",
            BodyTemplate = "Beste {{Voornaam}}, kan {{Datumopties}}?",
            PlanningFields = PlanningFields.Build(new DateTime(2026, 11, 1), new[] { new DateTime(2026, 11, 4) }),
            DelayBetweenMails = TimeSpan.Zero,
        };

        await BulkMailSender.SendAsync(sender, new[] { Ontvanger("Anne", "anne@school.nl") }, opties,
            cancellationToken: TestContext.Current.CancellationToken);

        Assert.Equal("Gastles in november", sender.Verstuurd[0].Subject);
        Assert.Equal("Beste Anne, kan woensdag 4 november?", sender.Verstuurd[0].Body);
    }

    [Fact]
    public async Task SendAsync_ZetOpmaakOmNaarHtmlWanneerGevraagd()
    {
        var sender = new FakeMailSender();
        var opties = new BulkSendOptions
        {
            SubjectTemplate = "Gastles",
            BodyTemplate = "Dit is **belangrijk**",
            IsHtml = true,
            DelayBetweenMails = TimeSpan.Zero,
        };

        await BulkMailSender.SendAsync(sender, new[] { Ontvanger("Anne", "anne@school.nl") }, opties,
            cancellationToken: TestContext.Current.CancellationToken);

        Assert.Contains("<b>belangrijk</b>", sender.Verstuurd[0].Body);
    }
}
