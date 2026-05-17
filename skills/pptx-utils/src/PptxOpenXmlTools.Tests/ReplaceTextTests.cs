using System.Text.Json;
using Xunit;

namespace PptxOpenXmlTools.Tests;

public sealed class ReplaceTextTests
{
    [Fact]
    public async Task ReplaceText_ReplacesExactTextAndOutputValidates()
    {
        var deck = PptxFixture.CreateSingleSlideDeck("Old text");
        var plan = Path.Combine(Path.GetTempPath(), $"replace-{Guid.NewGuid():N}.json");
        var output = Path.Combine(Path.GetTempPath(), $"out-{Guid.NewGuid():N}.pptx");
        await File.WriteAllTextAsync(plan, """
        {
          "replacements": [
            {
              "slideIndex": 1,
              "shapeIndex": 1,
              "find": "Old text",
              "replace": "New text",
              "mode": "exact"
            }
          ]
        }
        """);

        var result = await Cli.RunAsync("replace-text", deck, "--plan", plan, "--output", output);

        Assert.Equal(0, result.ExitCode);
        Assert.True(File.Exists(output));

        var textMap = await Cli.RunAsync("text-map", output, "--format", "json");
        using var doc = JsonDocument.Parse(textMap.Stdout);
        Assert.Equal("New text", doc.RootElement.GetProperty("items")[0].GetProperty("text").GetString());

        var validate = await Cli.RunAsync("validate", output, "--format", "json");
        Assert.Equal(0, validate.ExitCode);
    }

    [Fact]
    public async Task ReplaceText_WhenTextDoesNotMatch_DoesNotLeaveOutput()
    {
        var deck = PptxFixture.CreateSingleSlideDeck("Actual text");
        var plan = Path.Combine(Path.GetTempPath(), $"replace-{Guid.NewGuid():N}.json");
        var output = Path.Combine(Path.GetTempPath(), $"out-{Guid.NewGuid():N}.pptx");
        await File.WriteAllTextAsync(plan, """
        {
          "replacements": [
            {
              "slideIndex": 1,
              "shapeIndex": 1,
              "find": "Expected text",
              "replace": "New text",
              "mode": "exact"
            }
          ]
        }
        """);

        var result = await Cli.RunAsync("replace-text", deck, "--plan", plan, "--output", output);

        Assert.Equal(1, result.ExitCode);
        Assert.False(File.Exists(output));
        Assert.Contains("Text mismatch", result.Stderr);
    }

    [Fact]
    public async Task ReplaceText_UsesPresentationSlideOrder()
    {
        var deck = PptxFixture.CreateDeckWithSlidePartInsertionOrderDifferentFromPresentationOrder();
        var plan = Path.Combine(Path.GetTempPath(), $"replace-{Guid.NewGuid():N}.json");
        var output = Path.Combine(Path.GetTempPath(), $"out-{Guid.NewGuid():N}.pptx");
        await File.WriteAllTextAsync(plan, """
        {
          "replacements": [
            {
              "slideIndex": 1,
              "shapeIndex": 1,
              "find": "First logical slide",
              "replace": "Updated first slide",
              "mode": "exact"
            }
          ]
        }
        """);

        var result = await Cli.RunAsync("replace-text", deck, "--plan", plan, "--output", output);

        Assert.Equal(0, result.ExitCode);
        var textMap = await Cli.RunAsync("text-map", output, "--format", "json");
        using var doc = JsonDocument.Parse(textMap.Stdout);
        Assert.Equal("Updated first slide", doc.RootElement.GetProperty("items")[0].GetProperty("text").GetString());
        Assert.Equal("Second logical slide", doc.RootElement.GetProperty("items")[1].GetProperty("text").GetString());
    }
}
