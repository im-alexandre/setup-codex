using System.Diagnostics;
using Xunit;

namespace DocxOpenXmlTools.Tests;

public sealed class DocxOpenXmlToolsCliTests
{
    [Fact]
    public async Task Validate_reports_no_openxml_errors_for_minimal_generated_docx()
    {
        var skillRoot = FindSkillRoot();
        var articleProject = Path.Combine(skillRoot, "src", "ArticleDocxBuilder", "ArticleDocxBuilder.csproj");
        var toolsProject = Path.Combine(skillRoot, "src", "DocxOpenXmlTools", "DocxOpenXmlTools.csproj");

        using var tempDir = new TempDirectory();
        var specPath = Path.Combine(tempDir.Path, "article_spec.json");
        var docxPath = Path.Combine(tempDir.Path, "article.docx");

        await File.WriteAllTextAsync(specPath, """
        {
          "title": "Teste Docx Utils",
          "subtitle": "Baseline automatizado",
          "authorLine": "Ultron",
          "resumo": "Documento minimo para validar Open XML.",
          "abstract": "Minimal document for Open XML validation.",
          "palavrasChave": ["teste"],
          "keywords": ["test"],
          "sections": [
            { "heading": "Introducao", "level": 1, "paragraphs": ["Paragrafo temporario de validacao."], "items": [] }
          ],
          "references": []
        }
        """);

        var buildResult = await RunProcessAsync("dotnet", $"run --project \"{articleProject}\" -- \"{specPath}\" \"{docxPath}\" Ultron", skillRoot);
        Assert.True(buildResult.ExitCode == 0, buildResult.CombinedOutput);
        Assert.True(File.Exists(docxPath), "O ArticleDocxBuilder nao gerou o DOCX temporario.");

        var validateResult = await RunProcessAsync("dotnet", $"run --project \"{toolsProject}\" -- validate \"{docxPath}\"", skillRoot);
        Assert.True(validateResult.ExitCode == 0, validateResult.CombinedOutput);
        Assert.Contains("TrackRevisions: True", validateResult.CombinedOutput);
        Assert.Contains("OpenXmlValidationErrors: 0", validateResult.CombinedOutput);
        Assert.Contains("OpenXmlValidationErrorsActionable: 0", validateResult.CombinedOutput);
    }

    private static string FindSkillRoot()
    {
        var directory = new DirectoryInfo(AppContext.BaseDirectory);
        while (directory is not null)
        {
            if (File.Exists(Path.Combine(directory.FullName, "scripts", "install-docx-utils.ps1")) &&
                Directory.Exists(Path.Combine(directory.FullName, "src", "DocxOpenXmlTools")))
            {
                return directory.FullName;
            }

            directory = directory.Parent;
        }

        throw new InvalidOperationException("Nao foi possivel localizar a raiz da skill docx-utils.");
    }

    private static async Task<ProcessResult> RunProcessAsync(string fileName, string arguments, string workingDirectory)
    {
        using var process = new Process();
        process.StartInfo = new ProcessStartInfo
        {
            FileName = fileName,
            Arguments = arguments,
            WorkingDirectory = workingDirectory,
            RedirectStandardOutput = true,
            RedirectStandardError = true,
            UseShellExecute = false,
            CreateNoWindow = true
        };

        process.Start();
        var stdout = await process.StandardOutput.ReadToEndAsync();
        var stderr = await process.StandardError.ReadToEndAsync();
        await process.WaitForExitAsync();

        return new ProcessResult(process.ExitCode, stdout, stderr);
    }

    private sealed record ProcessResult(int ExitCode, string StandardOutput, string StandardError)
    {
        public string CombinedOutput => StandardOutput + Environment.NewLine + StandardError;
    }

    private sealed class TempDirectory : IDisposable
    {
        public TempDirectory()
        {
            Path = System.IO.Path.Combine(System.IO.Path.GetTempPath(), "docx-utils-test-" + Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(Path);
        }

        public string Path { get; }

        public void Dispose()
        {
            if (Directory.Exists(Path))
            {
                Directory.Delete(Path, recursive: true);
            }
        }
    }
}
