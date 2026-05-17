using System.Text.Json;
using Xunit;

namespace PptxOpenXmlTools.Tests;

public sealed class InspectTests
{
    [Fact]
    public async Task Inspect_ReturnsSlideAndShapeCounts()
    {
        var deck = PptxFixture.CreateSingleSlideDeck("Hello");

        var result = await Cli.RunAsync("inspect", deck, "--format", "json");

        Assert.Equal(0, result.ExitCode);
        using var doc = JsonDocument.Parse(result.Stdout);
        Assert.Equal(1, doc.RootElement.GetProperty("slideCount").GetInt32());
        Assert.Equal(1, doc.RootElement.GetProperty("slides")[0].GetProperty("textShapeCount").GetInt32());
    }
}
