using System.Text.Json;
using Xunit;

namespace PptxOpenXmlTools.Tests;

public sealed class TextMapTests
{
    [Fact]
    public async Task TextMap_ReturnsEditableTextWithSlideAndShapeIndexes()
    {
        var deck = PptxFixture.CreateSingleSlideDeck("Hello");

        var result = await Cli.RunAsync("text-map", deck, "--format", "json");

        Assert.Equal(0, result.ExitCode);
        using var doc = JsonDocument.Parse(result.Stdout);
        var item = doc.RootElement.GetProperty("items")[0];
        Assert.Equal(1, item.GetProperty("slideIndex").GetInt32());
        Assert.Equal(1, item.GetProperty("shapeIndex").GetInt32());
        Assert.Equal("Hello", item.GetProperty("text").GetString());
    }

    [Fact]
    public async Task TextMap_UsesPresentationSlideOrder()
    {
        var deck = PptxFixture.CreateDeckWithSlidePartInsertionOrderDifferentFromPresentationOrder();

        var result = await Cli.RunAsync("text-map", deck, "--format", "json");

        Assert.Equal(0, result.ExitCode);
        using var doc = JsonDocument.Parse(result.Stdout);
        Assert.Equal("First logical slide", doc.RootElement.GetProperty("items")[0].GetProperty("text").GetString());
        Assert.Equal("Second logical slide", doc.RootElement.GetProperty("items")[1].GetProperty("text").GetString());
    }
}
