using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Xunit;

namespace DocxOpenXmlTools.Tests;

public sealed class TemplateApplicationAuditCliTests
{
    [Fact]
    public async Task Audit_template_application_produces_markdown_report_and_flags_missing_required_regions()
    {
        using var tempDir = TemplateCliTestSupport.CreateTempDirectory();
        var templatePath = TemplateCliTestSupport.CreateSyntheticTemplateDocx(tempDir.Path);
        var sourcePath = TemplateCliTestSupport.CreateSyntheticSourceDocx(tempDir.Path);
        var profilePath = TemplateCliTestSupport.WriteCanonicalProfile(tempDir.Path, templatePath);
        var outputPath = Path.Combine(tempDir.Path, "audited.docx");
        var reportPath = Path.Combine(tempDir.Path, "audit.md");

        var applyResult = await TemplateCliTestSupport.RunToolsAsync(
            $"apply-template --template \"{templatePath}\" --source \"{sourcePath}\" --profile \"{profilePath}\" --out \"{outputPath}\"");

        Assert.Equal(0, applyResult.ExitCode);

        var auditResult = await TemplateCliTestSupport.RunToolsAsync(
            $"audit-template-application \"{outputPath}\" --profile \"{profilePath}\" --report \"{reportPath}\"");

        Assert.NotEqual(0, auditResult.ExitCode);
        Assert.True(File.Exists(reportPath), "O relatorio Markdown de auditoria nao foi gerado.");

        var reportText = await File.ReadAllTextAsync(reportPath);
        Assert.Contains("# Template Application Audit", reportText);
        Assert.Contains("pendencias", reportText, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("Open XML", reportText, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public async Task Audit_template_application_detects_unfilled_mandatory_region_in_the_source_document()
    {
        using var tempDir = TemplateCliTestSupport.CreateTempDirectory();
        var templatePath = TemplateCliTestSupport.CreateSyntheticTemplateDocx(tempDir.Path);
        var sourcePath = TemplateCliTestSupport.CreateSyntheticSourceDocx(tempDir.Path);
        var profilePath = TemplateCliTestSupport.WriteCanonicalProfile(tempDir.Path, templatePath);
        var outputPath = Path.Combine(tempDir.Path, "audited-missing-region.docx");

        var applyResult = await TemplateCliTestSupport.RunToolsAsync(
            $"apply-template --template \"{templatePath}\" --source \"{sourcePath}\" --profile \"{profilePath}\" --out \"{outputPath}\"");

        Assert.Equal(0, applyResult.ExitCode);

        ReinstateCanonicalMarkerOutsideReferences(outputPath);

        var auditResult = await TemplateCliTestSupport.RunToolsAsync(
            $"audit-template-application \"{outputPath}\" --profile \"{profilePath}\"");

        Assert.NotEqual(0, auditResult.ExitCode);
        Assert.Contains("Regiao obrigatoria pendente", auditResult.CombinedOutput, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("REGIAO PRESERVADA DO TEMPLATE", auditResult.CombinedOutput, StringComparison.OrdinalIgnoreCase);
    }

    private static void ReinstateCanonicalMarkerOutsideReferences(string docxPath)
    {
        using var document = WordprocessingDocument.Open(docxPath, true);
        var body = document.MainDocumentPart!.Document.Body!;
        body.Append(new Paragraph(new Run(new Text("REGIAO PRESERVADA DO TEMPLATE"))));
        document.MainDocumentPart.Document.Save();
    }
}
