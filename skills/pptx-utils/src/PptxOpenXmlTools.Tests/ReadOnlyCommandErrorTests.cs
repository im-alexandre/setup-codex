using System.Text.Json;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using System.IO.Compression;
using System.Xml.Linq;
using Xunit;

namespace PptxOpenXmlTools.Tests;

public sealed class ReadOnlyCommandErrorTests
{
    [Fact]
    public async Task ReadOnlyCommands_WhenPptxIsCorrupted_ReturnControlledError()
    {
        var corrupted = await WriteCorruptedPptxAsync();

        foreach (var command in new[] { "inspect", "text-map", "validate" })
        {
            var result = await Cli.RunAsync(command, corrupted, "--format", "json");

            Assert.Equal(2, result.ExitCode);
            Assert.Contains("Invalid PPTX", result.Stderr);
            Assert.DoesNotContain("System.", result.Stderr);
        }
    }

    [Fact]
    public async Task Validate_WhenSlideIdListIsMissing_ReturnsInvalidDocument()
    {
        var pptx = CreatePresentationWithoutSlideIdList();

        var result = await Cli.RunAsync("validate", pptx, "--format", "json");

        Assert.Equal(1, result.ExitCode);
        using var doc = JsonDocument.Parse(result.Stdout);
        Assert.False(doc.RootElement.GetProperty("isValid").GetBoolean());
        Assert.True(doc.RootElement.GetProperty("hasPresentation").GetBoolean());
        Assert.False(doc.RootElement.GetProperty("hasSlideIdList").GetBoolean());
        Assert.Equal(0, doc.RootElement.GetProperty("openXmlValidationErrors").GetInt32());
    }

    private static async Task<string> WriteCorruptedPptxAsync()
    {
        var path = Path.Combine(Path.GetTempPath(), $"pptx-utils-{Guid.NewGuid():N}.pptx");
        await File.WriteAllTextAsync(path, "not a pptx");
        return path;
    }

    private static string CreatePresentationWithoutSlideIdList()
    {
        var path = PptxFixture.CreateSingleSlideDeck("Text");
        using var archive = ZipFile.Open(path, ZipArchiveMode.Update);
        var entry = archive.GetEntry("ppt/presentation.xml")
            ?? throw new InvalidOperationException("presentation.xml not found");
        XDocument document;
        using (var stream = entry.Open())
        {
            document = XDocument.Load(stream);
        }

        XNamespace p = "http://schemas.openxmlformats.org/presentationml/2006/main";
        document.Root?.Element(p + "sldIdLst")?.Remove();

        entry.Delete();
        var rewritten = archive.CreateEntry("ppt/presentation.xml");
        using var rewrittenStream = rewritten.Open();
        document.Save(rewrittenStream);
        return path;
    }
}
