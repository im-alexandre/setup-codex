using System.IO.Compression;
using System.Text.Json;
using Xunit;

namespace DocxOpenXmlTools.Tests;

public sealed class TemplateApplicationCliTests
{
    [Fact]
    public async Task Apply_template_preserves_undeclared_regions_and_emits_an_application_document()
    {
        using var tempDir = TemplateCliTestSupport.CreateTempDirectory();
        var templatePath = TemplateCliTestSupport.CreateSyntheticTemplateDocx(tempDir.Path);
        var sourcePath = TemplateCliTestSupport.CreateSyntheticSourceDocx(tempDir.Path);
        var profilePath = TemplateCliTestSupport.WriteCanonicalProfile(tempDir.Path, templatePath);
        var outputPath = Path.Combine(tempDir.Path, "aplicado.docx");

        var result = await TemplateCliTestSupport.RunToolsAsync(
            $"apply-template --template \"{templatePath}\" --source \"{sourcePath}\" --profile \"{profilePath}\" --out \"{outputPath}\" --report \"{Path.Combine(tempDir.Path, "application-report.md")}\"");

        Assert.Equal(0, result.ExitCode);
        Assert.True(File.Exists(outputPath), "O DOCX aplicado nao foi gerado.");

        using var zip = ZipFile.OpenRead(outputPath);
        Assert.Contains(zip.Entries, entry => entry.FullName == "word/document.xml");

        var report = Path.Combine(tempDir.Path, "application-report.md");
        Assert.True(File.Exists(report), "O relatorio de aplicacao nao foi gerado.");

        var reportText = await File.ReadAllTextAsync(report);
        Assert.Contains("# Template Application Report", reportText);
        Assert.Contains("Regioes preservadas", reportText);
        Assert.Contains("Regioes aplicadas", reportText);
        Assert.Contains("REGIAO PRESERVADA DO TEMPLATE", reportText, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public async Task Apply_template_rejects_profile_when_template_hash_diverges()
    {
        using var tempDir = TemplateCliTestSupport.CreateTempDirectory();
        var templatePath = TemplateCliTestSupport.CreateSyntheticTemplateDocx(tempDir.Path);
        var sourcePath = TemplateCliTestSupport.CreateSyntheticSourceDocx(tempDir.Path);
        var profilePath = TemplateCliTestSupport.WriteCanonicalProfile(tempDir.Path, templatePath);
        var outputPath = Path.Combine(tempDir.Path, "aplicado-hash-divergente.docx");

        TemplateCliTestSupport.MutateTemplateBodyForHashMismatch(templatePath);

        var result = await TemplateCliTestSupport.RunToolsAsync(
            $"apply-template --template \"{templatePath}\" --source \"{sourcePath}\" --profile \"{profilePath}\" --out \"{outputPath}\"");

        Assert.NotEqual(0, result.ExitCode);
        Assert.Contains("hash", result.CombinedOutput, StringComparison.OrdinalIgnoreCase);
        Assert.False(File.Exists(outputPath), "O output nao deveria ser produzido com hash divergente.");
    }
}
