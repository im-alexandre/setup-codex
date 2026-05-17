using System.Text.Json;
using Xunit;

namespace PptxOpenXmlTools.Tests;

public sealed class ValidateTests
{
    [Fact]
    public async Task Validate_ReturnsJsonWithZeroOpenXmlErrorsForFixture()
    {
        var deck = PptxFixture.CreateSingleSlideDeck("Hello");

        var result = await Cli.RunAsync("validate", deck, "--format", "json");

        Assert.Equal(0, result.ExitCode);
        using var doc = JsonDocument.Parse(result.Stdout);
        Assert.True(doc.RootElement.GetProperty("isValid").GetBoolean());
        Assert.True(doc.RootElement.GetProperty("hasPresentation").GetBoolean());
        Assert.True(doc.RootElement.GetProperty("hasSlideIdList").GetBoolean());
        Assert.Equal(0, doc.RootElement.GetProperty("openXmlValidationErrors").GetInt32());
    }
}
