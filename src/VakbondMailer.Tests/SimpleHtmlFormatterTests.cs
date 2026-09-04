using VakbondMailer.Services;
using Xunit;

namespace VakbondMailer.Tests;

public class SimpleHtmlFormatterTests
{
    [Fact]
    public void ToHtml_ConvertsBoldMarkdown()
    {
        var html = SimpleHtmlFormatter.ToHtml("Dit is **belangrijk** nieuws.");

        Assert.Contains("<b>belangrijk</b>", html);
    }

    [Fact]
    public void ToHtml_ConvertsItalicMarkdown()
    {
        var html = SimpleHtmlFormatter.ToHtml("Dit is *nadrukkelijk* zo.");

        Assert.Contains("<i>nadrukkelijk</i>", html);
    }

    [Fact]
    public void ToHtml_LinkifiesBareUrls()
    {
        var html = SimpleHtmlFormatter.ToHtml("Zie https://www.fnv.nl voor meer info.");

        Assert.Contains("<a href=\"https://www.fnv.nl\">https://www.fnv.nl</a>", html);
    }

    [Fact]
    public void ToHtml_EscapesHtmlSpecialCharacters()
    {
        var html = SimpleHtmlFormatter.ToHtml("Salaris < 3000 & > 2000");

        Assert.DoesNotContain("< 3000", html);
        Assert.Contains("&lt; 3000", html);
        Assert.Contains("&amp;", html);
    }

    [Fact]
    public void ToHtml_ConvertsNewlinesToBreaks()
    {
        var html = SimpleHtmlFormatter.ToHtml("Regel een\nRegel twee");

        Assert.Contains("Regel een<br>", html);
    }
}
