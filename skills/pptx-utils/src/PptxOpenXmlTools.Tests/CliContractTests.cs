using Xunit;

namespace PptxOpenXmlTools.Tests;

public sealed class CliContractTests
{
    [Fact]
    public async Task NoArguments_PrintsHelpAndExitsZero()
    {
        var result = await Cli.RunAsync();

        Assert.Equal(0, result.ExitCode);
        AssertHelpShell(result.Stdout);
        Assert.Empty(result.Stderr);
    }

    [Theory]
    [InlineData("--help")]
    [InlineData("-h")]
    [InlineData("help")]
    public async Task HelpAliases_PrintsTheSameHelpAndExitsZero(string helpArg)
    {
        var result = await Cli.RunAsync(helpArg);

        Assert.Equal(0, result.ExitCode);
        AssertHelpShell(result.Stdout);
        Assert.Empty(result.Stderr);
    }

    [Fact]
    public async Task Help_DoesNotTreatPlannedCommandsAsAvailable()
    {
        var result = await Cli.RunAsync("--help");

        var stdout = Normalize(result.Stdout);
        var plannedIndex = stdout.IndexOf("Planned:", StringComparison.Ordinal);

        Assert.True(plannedIndex >= 0, "Help output should include a Planned section.");

        var availableCommandsSection = stdout[..plannedIndex];

        Assert.Contains("Commands:", availableCommandsSection);
        Assert.Contains("inspect", availableCommandsSection);
        Assert.Contains("text-map", availableCommandsSection);
        Assert.Contains("replace-text", availableCommandsSection);
        Assert.Contains("validate", availableCommandsSection);
        Assert.DoesNotContain("duplicate-slide", availableCommandsSection);
        Assert.DoesNotContain("delete-slide", availableCommandsSection);
        Assert.DoesNotContain("reorder-slides", availableCommandsSection);
        Assert.DoesNotContain("replace-image", availableCommandsSection);
        Assert.DoesNotContain("render-preview", availableCommandsSection);
    }

    [Fact]
    public async Task PlannedCommand_ReturnsUnknownAndWritesError()
    {
        var result = await Cli.RunAsync("duplicate-slide");

        Assert.Equal(2, result.ExitCode);
        Assert.Empty(result.Stdout);
        Assert.Contains("Unknown command: duplicate-slide", result.Stderr);
    }

    [Theory]
    [InlineData("nonsense-command")]
    [InlineData("bogus")]
    public async Task UnknownCommands_WriteOnlyToStderr(string command)
    {
        var result = await Cli.RunAsync(command);

        Assert.Equal(2, result.ExitCode);
        Assert.True(string.IsNullOrWhiteSpace(result.Stdout));
        Assert.Contains($"Unknown command: {command}", result.Stderr);
    }

    private static void AssertHelpShell(string stdout)
    {
        var normalized = Normalize(stdout);

        Assert.Contains("pptx-utils", normalized);
        Assert.Contains("Commands:", normalized);
        Assert.Contains("Planned:", normalized);
        Assert.Contains("inspect", normalized);
        Assert.Contains("text-map", normalized);
        Assert.Contains("replace-text", normalized);
        Assert.Contains("validate", normalized);
        Assert.Contains("duplicate-slide", normalized);
    }

    private static string Normalize(string value) => value.Replace("\r\n", "\n");
}
