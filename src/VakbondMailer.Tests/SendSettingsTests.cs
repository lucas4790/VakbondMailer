using VakbondMailer.Services;
using Xunit;

namespace VakbondMailer.Tests;

public class SendSettingsTests
{
    [Theory]
    [InlineData("anne@fnv.nl")]
    [InlineData("Anne@FNV.NL")]
    [InlineData("anne.de.boer@fnv.nl")]
    public void IsAllowedSender_LaatFnvAdressenDoor(string email)
    {
        Assert.True(SendSettings.IsAllowedSender(email));
    }

    [Theory]
    [InlineData("anne@gmail.com")]
    [InlineData("")]
    [InlineData("   ")]
    [InlineData("anne@fnv.nl.example.com")] // bevat het domein wel, maar eindigt er niet op
    public void IsAllowedSender_WeigertDeRest(string email)
    {
        Assert.False(SendSettings.IsAllowedSender(email));
    }

    [Theory]
    [InlineData("2", 2)]
    [InlineData("1.5", 1.5)]
    [InlineData("1,5", 1.5)]    // Nederlandse Windows: komma als decimaalteken
    [InlineData("  3  ", 3)]
    public void ParseDelaySeconds_LeestWatErStaat(string invoer, double verwacht)
    {
        Assert.Equal(verwacht, SendSettings.ParseDelaySeconds(invoer));
    }

    [Theory]
    [InlineData("")]
    [InlineData("abc")]
    [InlineData("-5")]
    public void ParseDelaySeconds_ValtTerugOpEenSecondeBijOnbruikbareInvoer(string invoer)
    {
        Assert.Equal(1, SendSettings.ParseDelaySeconds(invoer));
    }

    [Theory]
    [InlineData("0")]
    [InlineData("0.05")]
    public void ParseDelaySeconds_HoudtEenMinimalePauzeAan(string invoer)
    {
        // Zonder pauze ziet Exchange een reeks mails al snel als spam.
        Assert.Equal(0.2, SendSettings.ParseDelaySeconds(invoer));
    }
}
