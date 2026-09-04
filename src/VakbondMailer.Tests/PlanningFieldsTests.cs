using VakbondMailer.Models;
using VakbondMailer.Services;
using Xunit;

namespace VakbondMailer.Tests;

public class PlanningFieldsTests
{
    [Fact]
    public void FormatMonthYear_UsesDutchMonthNames()
    {
        Assert.Equal("november 2026", PlanningFields.FormatMonthYear(new DateTime(2026, 11, 1)));
    }

    [Fact]
    public void FormatDateOptions_IsEmptyWhenNothingChosen()
    {
        Assert.Equal(string.Empty, PlanningFields.FormatDateOptions(Array.Empty<DateTime>()));
    }

    [Fact]
    public void FormatDateOptions_SingleDateHasNoConjunction()
    {
        var result = PlanningFields.FormatDateOptions(new[] { new DateTime(2026, 11, 4) });

        Assert.Equal("woensdag 4 november", result);
    }

    [Fact]
    public void FormatDateOptions_ReadsAsAProposalAndIsSortedByDate()
    {
        var result = PlanningFields.FormatDateOptions(new[]
        {
            new DateTime(2026, 11, 24),
            new DateTime(2026, 11, 4),
            new DateTime(2026, 11, 13),
        });

        Assert.Equal("woensdag 4 november, vrijdag 13 november of dinsdag 24 november", result);
    }

    [Fact]
    public void FormatDateOptions_IgnoresDuplicates()
    {
        var result = PlanningFields.FormatDateOptions(new[]
        {
            new DateTime(2026, 11, 4),
            new DateTime(2026, 11, 4),
        });

        Assert.Equal("woensdag 4 november", result);
    }

    [Fact]
    public void Render_FillsPlanningFieldsThatAreNotInTheList()
    {
        var recipient = new Recipient
        {
            Email = "docent@voorbeeldschool.nl",
            Fields = new Dictionary<string, string> { ["Voornaam"] = "Anne" },
        };
        var planning = PlanningFields.Build(new DateTime(2026, 11, 1), new[] { new DateTime(2026, 11, 4) });

        var result = TemplateRenderer.Render(
            "Beste {{Voornaam}}, kan {{Datumopties}} in {{Maand}}?", recipient, planning);

        Assert.Equal("Beste Anne, kan woensdag 4 november in november?", result);
    }

    [Fact]
    public void Render_ListColumnWinsOverPlanningFieldWithSameName()
    {
        var recipient = new Recipient
        {
            Email = "docent@voorbeeldschool.nl",
            Fields = new Dictionary<string, string> { ["Maand"] = "januari" },
        };
        var planning = PlanningFields.Build(new DateTime(2026, 11, 1), Array.Empty<DateTime>());

        var result = TemplateRenderer.Render("In {{Maand}}", recipient, planning);

        Assert.Equal("In januari", result);
    }

    [Fact]
    public void NextMonths_BegintBijDeMaandVanVandaag()
    {
        var maanden = PlanningFields.NextMonths(3, new DateTime(2026, 10, 17));

        Assert.Equal(new[]
        {
            new DateTime(2026, 10, 1),
            new DateTime(2026, 11, 1),
            new DateTime(2026, 12, 1),
        }, maanden);
    }

    [Fact]
    public void NextMonths_LooptNetjesDoorNaarHetVolgendeJaar()
    {
        var maanden = PlanningFields.NextMonths(4, new DateTime(2026, 11, 30));

        Assert.Equal(new DateTime(2027, 2, 1), maanden[3]);
    }
}
