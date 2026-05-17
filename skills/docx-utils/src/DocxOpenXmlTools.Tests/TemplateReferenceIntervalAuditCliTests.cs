using System.Text;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Xunit;

namespace DocxOpenXmlTools.Tests;

public sealed class TemplateReferenceIntervalAuditCliTests
{
    [Fact]
    public async Task Audit_template_application_passes_when_references_tail_is_clean_after_apply_template()
    {
        using var tempDir = TemplateCliTestSupport.CreateTempDirectory();
        var templatePath = TemplateCliTestSupport.CreateSyntheticTemplateDocx(
            tempDir.Path,
            preservedRegionText: "BLOCO 2 DO TEMPLATE");
        var sourcePath = TemplateCliTestSupport.CreateSyntheticSourceDocx(
            tempDir.Path,
            preservedRegionText: "BLOCO 2 DA ORIGEM");
        var profilePath = TemplateCliTestSupport.WriteCanonicalProfile(tempDir.Path, templatePath);
        var outputPath = Path.Combine(tempDir.Path, "aplicado-references-limpo.docx");
        var reportPath = Path.Combine(tempDir.Path, "audit-references-limpo.md");

        var applyResult = await TemplateCliTestSupport.RunToolsAsync(
            $"apply-template --template \"{templatePath}\" --source \"{sourcePath}\" --profile \"{profilePath}\" --out \"{outputPath}\"");

        Assert.Equal(0, applyResult.ExitCode);
        Assert.True(File.Exists(outputPath), "O DOCX aplicado nao foi gerado.");

        using (var document = WordprocessingDocument.Open(outputPath, false))
        {
            var tailParagraph = document.MainDocumentPart!.Document.Body!
                .Elements<Paragraph>()
                .Last();

            var tailText = string.Concat(tailParagraph.Descendants<Text>().Select(text => text.Text));
            Assert.True(string.IsNullOrWhiteSpace(tailText), "O tail de referencias deveria ficar vazio apos o apply-template.");
        }

        var auditResult = await TemplateCliTestSupport.RunToolsAsync(
            $"audit-template-application \"{outputPath}\" --profile \"{profilePath}\" --report \"{reportPath}\"");

        Assert.Equal(0, auditResult.ExitCode);
        Assert.True(File.Exists(reportPath), "O relatorio de auditoria deveria ser gerado.");

        var reportText = await File.ReadAllTextAsync(reportPath, Encoding.UTF8);
        Assert.Contains("Pendencias: 0", reportText, StringComparison.OrdinalIgnoreCase);
        Assert.DoesNotContain("p-0006", reportText, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public async Task Audit_template_application_fails_when_references_tail_keeps_custom_residue_after_apply_template()
    {
        using var tempDir = TemplateCliTestSupport.CreateTempDirectory();
        var templatePath = TemplateCliTestSupport.CreateSyntheticTemplateDocx(
            tempDir.Path,
            preservedRegionText: "BLOCO 2 DO TEMPLATE");
        var sourcePath = TemplateCliTestSupport.CreateSyntheticSourceDocx(
            tempDir.Path,
            preservedRegionText: "BLOCO 2 DA ORIGEM");
        var profilePath = TemplateCliTestSupport.WriteCanonicalProfile(tempDir.Path, templatePath);
        var outputPath = Path.Combine(tempDir.Path, "aplicado-references-sujo.docx");
        var reportPath = Path.Combine(tempDir.Path, "audit-references-sujo.md");

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
            tailParagraph.Append(new Run(new Text("RESIDUO CUSTOMIZADO APOS REFERENCES")));
            document.MainDocumentPart.Document.Save();
        }

        using (var document = WordprocessingDocument.Open(outputPath, false))
        {
            var tailParagraph = document.MainDocumentPart!.Document.Body!
                .Elements<Paragraph>()
                .Last();

            var tailText = string.Concat(tailParagraph.Descendants<Text>().Select(text => text.Text));
            Assert.Contains("RESIDUO CUSTOMIZADO", tailText, StringComparison.OrdinalIgnoreCase);
        }

        var auditResult = await TemplateCliTestSupport.RunToolsAsync(
            $"audit-template-application \"{outputPath}\" --profile \"{profilePath}\" --report \"{reportPath}\"");

        Assert.NotEqual(0, auditResult.ExitCode);
        Assert.True(File.Exists(reportPath), "O relatorio de auditoria deveria ser gerado.");

        var reportText = await File.ReadAllTextAsync(reportPath, Encoding.UTF8);
        Assert.Contains("Pendencias", reportText, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("p-0006", reportText, StringComparison.OrdinalIgnoreCase);
    }
}
