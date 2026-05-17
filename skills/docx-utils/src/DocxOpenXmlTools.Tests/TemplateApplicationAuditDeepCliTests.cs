using System.Text;
using System.Text.Json;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Xunit;

namespace DocxOpenXmlTools.Tests;

public sealed class TemplateApplicationAuditDeepCliTests
{
    [Fact]
    public async Task Audit_template_application_with_report_still_returns_nonzero_for_critical_pending_issues()
    {
        using var tempDir = TemplateCliTestSupport.CreateTempDirectory();
        var templatePath = TemplateCliTestSupport.CreateSyntheticTemplateDocx(tempDir.Path);
        var sourcePath = TemplateCliTestSupport.CreateSyntheticSourceDocx(tempDir.Path);
        var profilePath = TemplateCliTestSupport.WriteCanonicalProfile(tempDir.Path, templatePath);
        var outputPath = Path.Combine(tempDir.Path, "aplicado-auditoria.docx");
        var reportPath = Path.Combine(tempDir.Path, "audit-report.md");

        await TemplateCliTestSupport.RunToolsAsync(
            $"apply-template --template \"{templatePath}\" --source \"{sourcePath}\" --profile \"{profilePath}\" --out \"{outputPath}\"");

        InjectCriticalPendingIssue(outputPath);

        var auditResult = await TemplateCliTestSupport.RunToolsAsync(
            $"audit-template-application \"{outputPath}\" --profile \"{profilePath}\" --report \"{reportPath}\"");

        Assert.NotEqual(0, auditResult.ExitCode);
        Assert.True(File.Exists(reportPath), "O relatorio de auditoria deveria ser gerado.");

        var reportText = await File.ReadAllTextAsync(reportPath, Encoding.UTF8);
        Assert.Contains("pendencia", reportText, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("PENDENCIA CRITICA", reportText, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("REGIAO PRESERVADA DO TEMPLATE", reportText, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public async Task Audit_template_application_uses_profile_regions_instead_of_only_hardcoded_marker_strings()
    {
        using var tempDir = TemplateCliTestSupport.CreateTempDirectory();
        var templatePath = TemplateCliTestSupport.CreateSyntheticTemplateDocx(
            tempDir.Path,
            fileName: "template-custom.docx",
            includeReferenceParagraph: false,
            titleHeading: "CAPA",
            titleBody: "Modelo sem marcadores",
            abstractHeading: "SINOPSE",
            abstractBody: "Texto-base do modelo.",
            preservedRegionText: "BLOCO GUARDADO");
        var sourcePath = TemplateCliTestSupport.CreateSyntheticSourceDocx(
            tempDir.Path,
            fileName: "source-custom.docx",
            includeReferenceParagraph: false,
            titleHeading: "DADOS ORIGINAIS",
            titleBody: "Texto-base da origem.",
            abstractHeading: "SINOPSE",
            abstractBody: "Resumo-base da origem.",
            preservedRegionText: "TRECHO DE FONTE");
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
            "profile-custom.json");
        var outputPath = Path.Combine(tempDir.Path, "aplicado-custom.docx");
        var reportPath = Path.Combine(tempDir.Path, "audit-custom-report.md");

        await TemplateCliTestSupport.RunToolsAsync(
            $"apply-template --template \"{templatePath}\" --source \"{sourcePath}\" --profile \"{profilePath}\" --out \"{outputPath}\"");

        Assert.True(File.Exists(outputPath), "O DOCX aplicado deveria existir para a auditoria");

        var auditResult = await TemplateCliTestSupport.RunToolsAsync(
            $"audit-template-application \"{outputPath}\" --profile \"{profilePath}\" --report \"{reportPath}\"");

        Assert.NotEqual(0, auditResult.ExitCode);
        Assert.True(File.Exists(reportPath), "O relatorio deveria ser criado mesmo com pendencias profile-driven.");

        var reportText = await File.ReadAllTextAsync(reportPath, Encoding.UTF8);
        Assert.Contains("abstract", reportText, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("p-9999", reportText, StringComparison.OrdinalIgnoreCase);
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

    private static void InjectCriticalPendingIssue(string docxPath)
    {
        using var document = WordprocessingDocument.Open(docxPath, true);
        var body = document.MainDocumentPart!.Document.Body!;
        body.Append(new Paragraph(new Run(new Text("PENDENCIA CRITICA: REGIAO PRESERVADA DO TEMPLATE"))));
        document.MainDocumentPart.Document.Save();
    }
}
