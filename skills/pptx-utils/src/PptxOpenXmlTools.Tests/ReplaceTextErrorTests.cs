using Xunit;

namespace PptxOpenXmlTools.Tests;

public sealed class ReplaceTextErrorTests
{
    [Fact]
    public async Task ReplaceText_WhenInputPptxIsCorrupted_ReturnsUsageErrorAndDoesNotLeaveOutput()
    {
        var corruptedDeck = Path.Combine(Path.GetTempPath(), $"corrupted-{Guid.NewGuid():N}.pptx");
        await File.WriteAllTextAsync(corruptedDeck, "not a pptx");
        var plan = await WritePlanAsync("""
        {
          "replacements": [
            {
              "slideIndex": 1,
              "shapeIndex": 1,
              "find": "Text",
              "replace": "New text",
              "mode": "exact"
            }
          ]
        }
        """);
        var output = Path.Combine(Path.GetTempPath(), $"out-{Guid.NewGuid():N}.pptx");

        var result = await Cli.RunAsync("replace-text", corruptedDeck, "--plan", plan, "--output", output);

        Assert.Equal(2, result.ExitCode);
        Assert.False(File.Exists(output));
        Assert.Contains("Invalid PPTX", result.Stderr);
        Assert.DoesNotContain("System.", result.Stderr);
    }

    [Fact]
    public async Task ReplaceText_WhenInputPptxDoesNotExist_ReturnsUsageErrorAndDoesNotCreateOutput()
    {
        var missingDeck = Path.Combine(Path.GetTempPath(), $"missing-{Guid.NewGuid():N}.pptx");
        var plan = await WritePlanAsync("""
        {
          "replacements": [
            {
              "slideIndex": 1,
              "shapeIndex": 1,
              "find": "Text",
              "replace": "New text",
              "mode": "exact"
            }
          ]
        }
        """);
        var output = Path.Combine(Path.GetTempPath(), $"out-{Guid.NewGuid():N}.pptx");

        var result = await Cli.RunAsync("replace-text", missingDeck, "--plan", plan, "--output", output);

        Assert.Equal(2, result.ExitCode);
        Assert.False(File.Exists(output));
        Assert.Contains("File not found", result.Stderr);
    }

    [Fact]
    public async Task ReplaceText_WhenShapeIndexIsOutOfRange_ReturnsUsageErrorAndDoesNotLeaveOutput()
    {
        var deck = PptxFixture.CreateSingleSlideDeck("Text");
        var plan = await WritePlanAsync("""
        {
          "replacements": [
            {
              "slideIndex": 1,
              "shapeIndex": 99,
              "find": "Text",
              "replace": "New text",
              "mode": "exact"
            }
          ]
        }
        """);
        var output = Path.Combine(Path.GetTempPath(), $"out-{Guid.NewGuid():N}.pptx");

        var result = await Cli.RunAsync("replace-text", deck, "--plan", plan, "--output", output);

        Assert.Equal(2, result.ExitCode);
        Assert.False(File.Exists(output));
        Assert.Contains("Shape index out of range", result.Stderr);
    }

    [Fact]
    public async Task ReplaceText_WhenOutputParentDirectoryDoesNotExist_ReturnsUsageErrorAndDoesNotCreateOutput()
    {
        var deck = PptxFixture.CreateSingleSlideDeck("Text");
        var plan = await WritePlanAsync("""
        {
          "replacements": [
            {
              "slideIndex": 1,
              "shapeIndex": 1,
              "find": "Text",
              "replace": "New text",
              "mode": "exact"
            }
          ]
        }
        """);
        var outputDirectory = Path.Combine(Path.GetTempPath(), $"missing-dir-{Guid.NewGuid():N}");
        var output = Path.Combine(outputDirectory, "out.pptx");

        var result = await Cli.RunAsync("replace-text", deck, "--plan", plan, "--output", output);

        Assert.Equal(2, result.ExitCode);
        Assert.False(File.Exists(output));
        Assert.Contains("Output directory not found", result.Stderr);
    }

    [Fact]
    public async Task ReplaceText_WhenPlanHasNoReplacements_ReturnsUsageErrorAndDoesNotCreateOutput()
    {
        var deck = PptxFixture.CreateSingleSlideDeck("Text");
        var plan = await WritePlanAsync("""
        {
          "replacements": []
        }
        """);
        var output = Path.Combine(Path.GetTempPath(), $"out-{Guid.NewGuid():N}.pptx");

        var result = await Cli.RunAsync("replace-text", deck, "--plan", plan, "--output", output);

        Assert.Equal(2, result.ExitCode);
        Assert.False(File.Exists(output));
        Assert.Contains("No replacements", result.Stderr);
    }

    [Fact]
    public async Task ReplaceText_WhenPlanOmitsReplacements_ReturnsUsageErrorAndDoesNotCreateOutput()
    {
        var deck = PptxFixture.CreateSingleSlideDeck("Text");
        var plan = await WritePlanAsync("{}");
        var output = Path.Combine(Path.GetTempPath(), $"out-{Guid.NewGuid():N}.pptx");

        var result = await Cli.RunAsync("replace-text", deck, "--plan", plan, "--output", output);

        Assert.Equal(2, result.ExitCode);
        Assert.False(File.Exists(output));
        Assert.Contains("No replacements", result.Stderr);
    }

    [Theory]
    [InlineData("""
        {
          "replacements": [
            {
              "shapeIndex": 1,
              "find": "Text",
              "replace": "New text",
              "mode": "exact"
            }
          ]
        }
        """, "slideIndex")]
    [InlineData("""
        {
          "replacements": [
            {
              "slideIndex": 1,
              "find": "Text",
              "replace": "New text",
              "mode": "exact"
            }
          ]
        }
        """, "shapeIndex")]
    [InlineData("""
        {
          "replacements": [
            {
              "slideIndex": 1,
              "shapeIndex": 1,
              "replace": "New text",
              "mode": "exact"
            }
          ]
        }
        """, "find")]
    [InlineData("""
        {
          "replacements": [
            {
              "slideIndex": 1,
              "shapeIndex": 1,
              "find": "Text",
              "mode": "exact"
            }
          ]
        }
        """, "replace")]
    public async Task ReplaceText_WhenRequiredPlanFieldsAreMissing_ReturnsUsageErrorAndDoesNotCreateOutput(
        string json,
        string missingField)
    {
        var deck = PptxFixture.CreateSingleSlideDeck("Text");
        var plan = await WritePlanAsync(json);
        var output = Path.Combine(Path.GetTempPath(), $"out-{Guid.NewGuid():N}.pptx");

        var result = await Cli.RunAsync("replace-text", deck, "--plan", plan, "--output", output);

        Assert.Equal(2, result.ExitCode);
        Assert.False(File.Exists(output));
        Assert.Contains(missingField, result.Stderr);
    }

    private static async Task<string> WritePlanAsync(string json)
    {
        var plan = Path.Combine(Path.GetTempPath(), $"replace-{Guid.NewGuid():N}.json");
        await File.WriteAllTextAsync(plan, json);
        return plan;
    }
}
