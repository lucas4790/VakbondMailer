using VakbondMailer.Services;
using Xunit;

namespace VakbondMailer.Tests;

public class SendHistoryServiceTests : IDisposable
{
    private readonly string _path = Path.Combine(Path.GetTempPath(), $"{Guid.NewGuid()}", "geschiedenis.json");

    public void Dispose()
    {
        var directory = Path.GetDirectoryName(_path);
        if (directory is not null && Directory.Exists(directory))
            Directory.Delete(directory, recursive: true);
    }

    [Fact]
    public void CountRecentlySent_IsZeroWithoutHistory()
    {
        var count = SendHistoryService.CountRecentlySent(
            _path, "Gastles", new[] { "a@school.nl" }, TimeSpan.FromDays(14), DateTime.Now);

        Assert.Equal(0, count);
    }

    [Fact]
    public void CountRecentlySent_CountsOnlyTheOverlapWithThisMailing()
    {
        var now = new DateTime(2026, 10, 1);
        SendHistoryService.Append(_path, "Gastles", new[] { "a@school.nl", "b@school.nl" }, now.AddDays(-3));

        var count = SendHistoryService.CountRecentlySent(
            _path, "Gastles", new[] { "a@school.nl", "b@school.nl", "c@school.nl" }, TimeSpan.FromDays(14), now);

        Assert.Equal(2, count);
    }

    [Fact]
    public void CountRecentlySent_IgnoresADifferentMailing()
    {
        var now = new DateTime(2026, 10, 1);
        SendHistoryService.Append(_path, "Gastles", new[] { "a@school.nl" }, now.AddDays(-1));

        var count = SendHistoryService.CountRecentlySent(
            _path, "Uitnodiging ledenvergadering", new[] { "a@school.nl" }, TimeSpan.FromDays(14), now);

        Assert.Equal(0, count);
    }

    [Fact]
    public void CountRecentlySent_IgnoresSendsOutsideTheWindow()
    {
        var now = new DateTime(2026, 10, 1);
        SendHistoryService.Append(_path, "Gastles", new[] { "a@school.nl" }, now.AddDays(-30));

        var count = SendHistoryService.CountRecentlySent(
            _path, "Gastles", new[] { "a@school.nl" }, TimeSpan.FromDays(14), now);

        Assert.Equal(0, count);
    }

    [Fact]
    public void CountRecentlySent_MatchesRegardlessOfCasingOrSpacing()
    {
        var now = new DateTime(2026, 10, 1);
        SendHistoryService.Append(_path, "Gastles", new[] { "Anne@School.nl" }, now.AddDays(-1));

        var count = SendHistoryService.CountRecentlySent(
            _path, "Gastles", new[] { " anne@school.nl " }, TimeSpan.FromDays(14), now);

        Assert.Equal(1, count);
    }

    [Fact]
    public void Append_DoesNotStoreReadableEmailAddresses()
    {
        SendHistoryService.Append(_path, "Gastles", new[] { "anne@school.nl" }, DateTime.Now);

        var contents = File.ReadAllText(_path);

        Assert.DoesNotContain("anne@school.nl", contents);
        Assert.Contains("Gastles", contents);
    }

    [Fact]
    public void Append_ForgetsEntriesOlderThanNinetyDays()
    {
        var now = new DateTime(2026, 10, 1);
        SendHistoryService.Append(_path, "Gastles", new[] { "oud@school.nl" }, now.AddDays(-200));
        SendHistoryService.Append(_path, "Gastles", new[] { "nieuw@school.nl" }, now);

        Assert.Single(SendHistoryService.Load(_path));
    }
}
