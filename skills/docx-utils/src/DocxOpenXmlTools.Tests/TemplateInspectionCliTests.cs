using System.Text.Json;
using Xunit;

namespace DocxOpenXmlTools.Tests;

public sealed class TemplateInspectionCliTests
{
    [Fact]
    public async Task Inspect_template_extracts_normal_style_direct_formatting_and_reference_runs()
    {
        using var tempDir = TemplateCliTestSupport.CreateTempDirectory();
        var templatePath = TemplateCliTestSupport.CreateSyntheticTemplateDocx(tempDir.Path);
        var candidatesPath = Path.Combine(tempDir.Path, "profile.candidates.json");
        var reportPath = Path.Combine(tempDir.Path, "profile.candidates.md");

        var result = await TemplateCliTestSupport.RunToolsAsync(
            $"inspect-template \"{templatePath}\" --out \"{candidatesPath}\" --report \"{reportPath}\"");

        Assert.Equal(0, result.ExitCode);
        Assert.True(File.Exists(candidatesPath), "O JSON de candidatos nao foi gerado.");
        Assert.True(File.Exists(reportPath), "O relatorio Markdown nao foi gerado.");

        using var json = JsonDocument.Parse(await File.ReadAllTextAsync(candidatesPath));
        Assert.Contains("template", json.RootElement.EnumerateObject().Select(p => p.Name));
        Assert.Contains("candidates", json.RootElement.EnumerateObject().Select(p => p.Name));

        var candidates = json.RootElement.GetProperty("candidates");
        Assert.NotEmpty(candidates.EnumerateArray());
        var joinedText = JsonSerializer.Serialize(json.RootElement);
        Assert.Contains("1 INTRODUCAO", joinedText);
        Assert.Contains("allCaps", joinedText);
        Assert.Contains("SILVA, J. A.", joinedText);
        Assert.Contains("Titulo do artigo.", joinedText);
        Assert.Contains("Revista X", joinedText);
        Assert.Contains("https://doi.org/10.1000/teste", joinedText);

        var report = await File.ReadAllTextAsync(reportPath);
        Assert.Contains("# Template Inspection Report", report);
        Assert.Contains("Normal", report);
        Assert.Contains("referencia", report, StringComparison.OrdinalIgnoreCase);
    }
}
