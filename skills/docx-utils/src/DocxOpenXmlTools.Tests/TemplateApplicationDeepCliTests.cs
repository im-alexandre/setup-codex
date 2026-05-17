using System.Text;
using System.Text.Json;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Xunit;

namespace DocxOpenXmlTools.Tests;

public sealed class TemplateApplicationDeepCliTests
{
    [Fact]
    public async Task Apply_template_returns_nonzero_when_required_regions_cannot_find_target_or_source()
    {
        using var tempDir = TemplateCliTestSupport.CreateTempDirectory();
        var templatePath = TemplateCliTestSupport.CreateSyntheticTemplateDocx(tempDir.Path);
        var sourcePath = TemplateCliTestSupport.CreateSyntheticSourceDocx(tempDir.Path, includeReferenceParagraph: false);
        var profilePath = WriteProfile(
            tempDir.Path,
            templatePath,
            new object[]
            {
                new
                {
                    role = "title",
                    templateBlockId = "p-0002",
                    replaceWith = "source.title",
                    preserveFormatting = true
                },
                new
                {
                    role = "abstract",
                    startBlockId = "p-9999",
                    endBlockId = "p-9999",
                    replaceWith = "source.abstract",
                    preserveFormattingFrom = "p-0003"
                },
                new
                {
                    role = "references",
                    startBlockId = "p-0005",
                    endBlockId = "p-0006",
                    replaceWith = "source.references",
                    referenceFormattingProfile = "refs-main"
                }
            },
            "profile-missing-regions.json");
        var outputPath = Path.Combine(tempDir.Path, "aplicado-missing-regions.docx");
        var reportPath = Path.Combine(tempDir.Path, "application-report.md");

        var result = await TemplateCliTestSupport.RunToolsAsync(
            $"apply-template --template \"{templatePath}\" --source \"{sourcePath}\" --profile \"{profilePath}\" --out \"{outputPath}\" --report \"{reportPath}\"");

        Assert.NotEqual(0, result.ExitCode);
        Assert.True(File.Exists(reportPath), "O relatorio deveria ser gerado mesmo com pendencias criticas.");

        var reportText = await File.ReadAllTextAsync(reportPath, Encoding.UTF8);
        Assert.Contains("No template block found for region `abstract`.", reportText, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("No source content found for region `references`.", reportText, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public async Task Apply_template_treats_start_and_end_block_ids_as_a_closed_interval()
    {
        using var tempDir = TemplateCliTestSupport.CreateTempDirectory();
        var templatePath = TemplateCliTestSupport.CreateSyntheticTemplateDocx(tempDir.Path);
        var sourcePath = TemplateCliTestSupport.CreateSyntheticSourceDocx(tempDir.Path);
        var profilePath = TemplateCliTestSupport.WriteCanonicalProfile(tempDir.Path, templatePath);
        var outputPath = Path.Combine(tempDir.Path, "aplicado-intervalo.docx");

        var result = await TemplateCliTestSupport.RunToolsAsync(
            $"apply-template --template \"{templatePath}\" --source \"{sourcePath}\" --profile \"{profilePath}\" --out \"{outputPath}\"");

        Assert.Equal(0, result.ExitCode);
        Assert.True(File.Exists(outputPath), "O DOCX aplicado nao foi gerado.");

        using var document = WordprocessingDocument.Open(outputPath, false);
        var paragraphTexts = document.MainDocumentPart!.Document.Body!
            .Elements<Paragraph>()
            .Select(p => string.Concat(p.Descendants<Text>().Select(t => t.Text)))
            .ToArray();

        Assert.Contains(paragraphTexts, text => text.Contains("Resumo de origem do documento.", StringComparison.OrdinalIgnoreCase));
        Assert.DoesNotContain(paragraphTexts, text => text.Contains("Conteudo exemplo do resumo.", StringComparison.OrdinalIgnoreCase));
    }

    [Fact]
    public async Task Apply_template_preserves_reference_runs_when_reference_formatting_profile_is_defined()
    {
        using var tempDir = TemplateCliTestSupport.CreateTempDirectory();
        var templatePath = TemplateCliTestSupport.CreateSyntheticTemplateDocx(tempDir.Path);
        var sourcePath = TemplateCliTestSupport.CreateSyntheticSourceDocx(tempDir.Path);
        var profilePath = TemplateCliTestSupport.WriteCanonicalProfile(tempDir.Path, templatePath);
        var outputPath = Path.Combine(tempDir.Path, "aplicado-referencias.docx");

        var result = await TemplateCliTestSupport.RunToolsAsync(
            $"apply-template --template \"{templatePath}\" --source \"{sourcePath}\" --profile \"{profilePath}\" --out \"{outputPath}\"");

        Assert.Equal(0, result.ExitCode);
        Assert.True(File.Exists(outputPath), "O DOCX aplicado nao foi gerado.");

        using var document = WordprocessingDocument.Open(outputPath, false);
        var referenceParagraph = document.MainDocumentPart!.Document.Body!
            .Elements<Paragraph>()
            .Single(paragraph =>
            {
                var text = string.Concat(paragraph.Descendants<Text>().Select(t => t.Text));
                return text.Contains("SILVA, J. A.", StringComparison.OrdinalIgnoreCase) &&
                       text.Contains("Titulo do artigo.", StringComparison.OrdinalIgnoreCase) &&
                       text.Contains("Revista X", StringComparison.OrdinalIgnoreCase);
            });

        var runs = referenceParagraph.Elements<Run>().ToArray();
        Assert.True(runs.Length >= 4, $"Esperava pelo menos 4 runs na referencia, mas encontrei {runs.Length}.");

        var authorRun = runs.SingleOrDefault(run => string.Concat(run.Descendants<Text>().Select(t => t.Text)).Contains("SILVA, J. A.", StringComparison.OrdinalIgnoreCase));
        Assert.NotNull(authorRun);
        Assert.NotNull(authorRun!.RunProperties?.Caps);

        var titleRun = runs.SingleOrDefault(run => string.Concat(run.Descendants<Text>().Select(t => t.Text)).Contains("Titulo do artigo.", StringComparison.OrdinalIgnoreCase));
        Assert.NotNull(titleRun);
        Assert.NotNull(titleRun!.RunProperties?.Bold);

        var publicationRun = runs.SingleOrDefault(run => string.Concat(run.Descendants<Text>().Select(t => t.Text)).Contains("Revista X", StringComparison.OrdinalIgnoreCase));
        Assert.NotNull(publicationRun);
        Assert.NotNull(publicationRun!.RunProperties?.Italic);

        var linkRun = runs.SingleOrDefault(run => string.Concat(run.Descendants<Text>().Select(t => t.Text)).Contains("https://doi.org/10.1000/teste", StringComparison.OrdinalIgnoreCase));
        Assert.NotNull(linkRun);
        var linkProperties = linkRun!.RunProperties;
        Assert.NotNull(linkProperties);
        Assert.NotNull(linkProperties.Color);
        Assert.Equal("0563C1", linkProperties.Color!.Val);
    }

    private static string WriteProfile(string directory, string templatePath, object[] regions, string profileFileName)
    {
        var profilePath = Path.Combine(directory, profileFileName);
        var profile = new
        {
            template = new
            {
                name = "template-profile-sintetico",
                sourceFile = Path.GetFileName(templatePath),
                sha256 = TemplateCliTestSupport.ComputeSha256(templatePath),
                profileVersion = 1
            },
            regions
        };

        File.WriteAllText(profilePath, JsonSerializer.Serialize(profile, new JsonSerializerOptions { WriteIndented = true }), Encoding.UTF8);
        return profilePath;
    }
}
