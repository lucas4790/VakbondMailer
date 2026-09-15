using VakbondMailer.Models;
using Xunit;

namespace VakbondMailer.Tests;

public class RecipientTests
{
    [Fact]
    public void Gelijkheid_HoudtRekeningMetDeFieldsDictionaryInstantie()
    {
        var a = new Recipient
        {
            Email = "anne@school.nl",
            Fields = new Dictionary<string, string> { ["Voornaam"] = "Anne" },
        };
        var b = new Recipient
        {
            Email = "anne@school.nl",
            Fields = new Dictionary<string, string> { ["Voornaam"] = "Anne" },
        };

        // Dictionary<TKey,TValue> heeft geen structurele Equals, dus twee losse instanties met
        // identieke inhoud zijn nooit aan elkaar gelijk. De record-conversie van Recipient
        // verandert dat niet: gelijkheid op Fields blijft op instantie-identiteit leunen.
        Assert.NotEqual(a, b);
    }

    [Fact]
    public void Gelijkheid_IsWaarBijDezelfdeFieldsInstantie()
    {
        var fields = new Dictionary<string, string> { ["Voornaam"] = "Anne" };
        var a = new Recipient { Email = "anne@school.nl", Fields = fields };
        var b = new Recipient { Email = "anne@school.nl", Fields = fields };

        Assert.Equal(a, b);
    }
}
