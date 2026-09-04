using VakbondMailer.Models;
using VakbondMailer.Services;
using Xunit;

namespace VakbondMailer.Tests;

public class RecipientSelectionTests
{
    private static ImportedRecipients Lijst(params string[] voornamen) => new()
    {
        Headers = new[] { "Voornaam", "E-mail" },
        Recipients = voornamen.Select(naam => new Recipient
        {
            Email = $"{naam.ToLowerInvariant()}@school.nl",
            Fields = new Dictionary<string, string>
            {
                ["Voornaam"] = naam,
                ["E-mail"] = $"{naam.ToLowerInvariant()}@school.nl",
            },
        }).ToList(),
        Warnings = Array.Empty<string>(),
    };

    [Fact]
    public void From_ZetIedereenStandaardAan()
    {
        var selectie = RecipientSelection.From(Lijst("Anne", "Bram", "Carla"));

        Assert.Equal(3, selectie.Selected.Count);
        Assert.Equal("3 ontvanger(s) geladen", selectie.CountLabel);
    }

    [Fact]
    public void Tabel_HeeftDeVinkjeskolomVoorDeGegevens()
    {
        var selectie = RecipientSelection.From(Lijst("Anne"));

        Assert.Equal(RecipientSelection.SelectionColumnName, selectie.Table.Columns[0].ColumnName);
        Assert.Equal("Voornaam", selectie.Table.Columns[1].ColumnName);
        Assert.Equal("Anne", selectie.Table.Rows[0]["Voornaam"]);
    }

    [Fact]
    public void RijIndexBlijftGekoppeldAanDezelfdeOntvanger()
    {
        var selectie = RecipientSelection.From(Lijst("Anne", "Bram", "Carla"));

        // Alleen de middelste rij uitvinken moet exact Bram overslaan.
        selectie.Table.Rows[1][RecipientSelection.SelectionColumnName] = false;

        Assert.Equal(new[] { "anne@school.nl", "carla@school.nl" }, selectie.Selected.Select(r => r.Email));
        Assert.Equal("Bram", selectie.At(1)!.Fields["Voornaam"]);
    }

    [Fact]
    public void SetAll_ZetAllesUitEnWeerAan()
    {
        var selectie = RecipientSelection.From(Lijst("Anne", "Bram"));

        selectie.SetAll(false);
        Assert.Empty(selectie.Selected);
        Assert.Equal("0 van 2 geselecteerd", selectie.CountLabel);

        selectie.SetAll(true);
        Assert.Equal(2, selectie.Selected.Count);
    }

    [Fact]
    public void SelectOnly_VinktPreciesDeMisluktenAan()
    {
        var lijst = Lijst("Anne", "Bram", "Carla");
        var selectie = RecipientSelection.From(lijst);
        var mislukt = new[] { lijst.Recipients[1] };

        selectie.SelectOnly(mislukt);

        Assert.Equal(new[] { "bram@school.nl" }, selectie.Selected.Select(r => r.Email));
        Assert.Equal("1 van 3 geselecteerd", selectie.CountLabel);
    }

    [Fact]
    public void At_ValtTerugOpDeEersteWanneerErNietsGeselecteerdIs()
    {
        var selectie = RecipientSelection.From(Lijst("Anne", "Bram"));

        // -1 is wat een DataGrid teruggeeft zonder selectie.
        Assert.Equal("anne@school.nl", selectie.At(-1)!.Email);
        Assert.Equal("anne@school.nl", selectie.At(99)!.Email);
    }

    [Fact]
    public void At_GeeftNullBijEenLegeLijst()
    {
        var selectie = RecipientSelection.From(Lijst());

        Assert.Null(selectie.At(0));
        Assert.Equal("0 ontvanger(s) geladen", selectie.CountLabel);
    }
}
