using Xunit;

namespace PptxOpenXmlTools.Tests;

public sealed class ErrorHandlingTests
{
    [Fact]
    public async Task UnknownCommand_ReturnsUsageError()
    {
        var result = await Cli.RunAsync("does-not-exist");

        Assert.Equal(2, result.ExitCode);
        Assert.Contains("Unknown command", result.Stderr);
    }

    [Theory]
    [InlineData("inspect")]
    [InlineData("text-map")]
    [InlineData("validate")]
    public async Task ReadOnlyCommands_WhenFileDoesNotExist_ReturnUsageError(string command)
    {
        var missing = Path.Combine(Path.GetTempPath(), $"missing-{Guid.NewGuid():N}.pptx");

        var result = await Cli.RunAsync(command, missing, "--format", "json");

        Assert.Equal(2, result.ExitCode);
        Assert.Contains("File not found", result.Stderr);
    }

    [Fact]
    public async Task ReplaceText_WhenPlanOptionIsMissing_ReturnsUsageErrorAndDoesNotCreateOutput()
    {
        var deck = PptxFixture.CreateSingleSlideDeck("Text");
        var output = Path.Combine(Path.GetTempPath(), $"out-{Guid.NewGuid():N}.pptx");

        var result = await Cli.RunAsync("replace-text", deck, "--output", output);

        Assert.Equal(2, result.ExitCode);
        Assert.False(File.Exists(output));
        Assert.Contains("Usage:", result.Stderr);
    }

    [Fact]
    public async Task ReplaceText_WhenPlanFileDoesNotExist_ReturnsUsageErrorAndDoesNotCreateOutput()
    {
        var deck = PptxFixture.CreateSingleSlideDeck("Text");
        var missingPlan = Path.Combine(Path.GetTempPath(), $"missing-{Guid.NewGuid():N}.json");
        var output = Path.Combine(Path.GetTempPath(), $"out-{Guid.NewGuid():N}.pptx");

        var result = await Cli.RunAsync("replace-text", deck, "--plan", missingPlan, "--output", output);

        Assert.Equal(2, result.ExitCode);
        Assert.False(File.Exists(output));
        Assert.Contains("Plan not found", result.Stderr);
    }

    [Fact]
    public async Task ReplaceText_WhenPlanJsonIsInvalid_ReturnsUsageErrorAndDoesNotCreateOutput()
    {
        var deck = PptxFixture.CreateSingleSlideDeck("Text");
        var plan = Path.Combine(Path.GetTempPath(), $"replace-{Guid.NewGuid():N}.json");
        var output = Path.Combine(Path.GetTempPath(), $"out-{Guid.NewGuid():N}.pptx");
        await File.WriteAllTextAsync(plan, "{ invalid json");

        var result = await Cli.RunAsync("replace-text", deck, "--plan", plan, "--output", output);

        Assert.Equal(2, result.ExitCode);
        Assert.False(File.Exists(output));
        Assert.Contains("Invalid replace-text plan", result.Stderr);
    }

    [Fact]
    public async Task ReplaceText_WhenModeIsUnsupported_ReturnsUsageErrorAndDoesNotLeaveOutput()
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
              "mode": "regex"
            }
          ]
        }
        """);
        var output = Path.Combine(Path.GetTempPath(), $"out-{Guid.NewGuid():N}.pptx");

        var result = await Cli.RunAsync("replace-text", deck, "--plan", plan, "--output", output);

        Assert.Equal(2, result.ExitCode);
        Assert.False(File.Exists(output));
        Assert.Contains("Unsupported mode", result.Stderr);
    }

    [Fact]
    public async Task ReplaceText_WhenSlideIndexIsOutOfRange_ReturnsUsageErrorAndDoesNotLeaveOutput()
    {
        var deck = PptxFixture.CreateSingleSlideDeck("Text");
        var plan = await WritePlanAsync("""
        {
          "replacements": [
            {
              "slideIndex": 99,
              "shapeIndex": 1,
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
        Assert.Contains("Slide index out of range", result.Stderr);
    }

    private static async Task<string> WritePlanAsync(string json)
    {
        var plan = Path.Combine(Path.GetTempPath(), $"replace-{Guid.NewGuid():N}.json");
        await File.WriteAllTextAsync(plan, json);
        return plan;
    }
}
