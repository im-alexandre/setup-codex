using System.Text;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Xunit;

namespace DocxOpenXmlTools.Tests;

public sealed class TemplateReferenceIntervalCliTests
{
    [Fact]
    public async Task Apply_template_replaces_the_entire_references_interval_instead_of_only_the_first_paragraph()
    {
        using var tempDir = TemplateCliTestSupport.CreateTempDirectory();
        var templatePath = TemplateCliTestSupport.CreateSyntheticTemplateDocx(
            tempDir.Path,
            preservedRegionText: "BLOCO 2 DO TEMPLATE");
        var sourcePath = TemplateCliTestSupport.CreateSyntheticSourceDocx(
            tempDir.Path,
            preservedRegionText: "BLOCO 2 DA ORIGEM");
        var profilePath = TemplateCliTestSupport.WriteCanonicalProfile(tempDir.Path, templatePath);
        var outputPath = Path.Combine(tempDir.Path, "aplicado-references-intervalo.docx");

        var result = await TemplateCliTestSupport.RunToolsAsync(
            $"apply-template --template \"{templatePath}\" --source \"{sourcePath}\" --profile \"{profilePath}\" --out \"{outputPath}\"");

        Assert.Equal(0, result.ExitCode);
        Assert.True(File.Exists(outputPath), "O DOCX aplicado nao foi gerado.");

        using var document = WordprocessingDocument.Open(outputPath, false);
        var paragraphTexts = document.MainDocumentPart!.Document.Body!
            .Elements<Paragraph>()
            .Select(paragraph => string.Concat(paragraph.Descendants<Text>().Select(text => text.Text)))
            .ToArray();

        Assert.Contains(paragraphTexts, text => text.Contains("SILVA, J. A.", StringComparison.OrdinalIgnoreCase));
        Assert.Contains(paragraphTexts, string.IsNullOrWhiteSpace);
        Assert.DoesNotContain(paragraphTexts, text => text.Contains("BLOCO 2 DO TEMPLATE", StringComparison.OrdinalIgnoreCase));
    }

    [Fact]
    public async Task Audit_template_application_fails_when_references_interval_keeps_custom_template_residue()
    {
        using var tempDir = TemplateCliTestSupport.CreateTempDirectory();
        var templatePath = TemplateCliTestSupport.CreateSyntheticTemplateDocx(
            tempDir.Path,
            preservedRegionText: "BLOCO 2 DO TEMPLATE");
        var sourcePath = TemplateCliTestSupport.CreateSyntheticSourceDocx(
            tempDir.Path,
            preservedRegionText: "BLOCO 2 DA ORIGEM");
        var profilePath = TemplateCliTestSupport.WriteCanonicalProfile(tempDir.Path, templatePath);
        var outputPath = Path.Combine(tempDir.Path, "aplicado-auditoria-referencias.docx");
        var reportPath = Path.Combine(tempDir.Path, "audit-referencias-report.md");

        var applyResult = await TemplateCliTestSupport.RunToolsAsync(
            $"apply-template --template \"{templatePath}\" --source \"{sourcePath}\" --profile \"{profilePath}\" --out \"{outputPath}\"");

        Assert.Equal(0, applyResult.ExitCode);
        Assert.True(File.Exists(outputPath), "O DOCX aplicado nao foi gerado.");

        using (var document = WordprocessingDocument.Open(outputPath, true))
        {
            var tailParagraph = document.MainDocumentPart!.Document.Body!
                .Elements<Paragraph>()
                .Last();

            tailParagraph.RemoveAllChildren<Run>();
            tailParagraph.Append(new Run(new Text("BLOCO 2 DO TEMPLATE")));
            document.MainDocumentPart.Document.Save();
        }

        var auditResult = await TemplateCliTestSupport.RunToolsAsync(
            $"audit-template-application \"{outputPath}\" --profile \"{profilePath}\" --report \"{reportPath}\"");

        Assert.NotEqual(0, auditResult.ExitCode);
        Assert.True(File.Exists(reportPath), "O relatorio de auditoria deveria ser gerado.");

        var reportText = await File.ReadAllTextAsync(reportPath, Encoding.UTF8);
        Assert.Contains("p-0006", reportText, StringComparison.OrdinalIgnoreCase);
    }
}
