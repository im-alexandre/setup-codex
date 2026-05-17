using System.Diagnostics;
using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace DocxOpenXmlTools.Tests;

internal static class TemplateCliTestSupport
{
    public static string FindSkillRoot()
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

    public static string GetToolsProjectPath()
    {
        return Path.Combine(FindSkillRoot(), "src", "DocxOpenXmlTools", "DocxOpenXmlTools.csproj");
    }

    public static async Task<ProcessResult> RunToolsAsync(string arguments, IReadOnlyDictionary<string, string?>? environment = null)
    {
        var skillRoot = FindSkillRoot();
        return await RunProcessAsync("dotnet", $"run --project \"{GetToolsProjectPath()}\" -- {arguments}", skillRoot, environment);
    }

    public static async Task<ProcessResult> RunProcessAsync(
        string fileName,
        string arguments,
        string workingDirectory,
        IReadOnlyDictionary<string, string?>? environment = null)
    {
        using var process = new Process();
        process.StartInfo = new ProcessStartInfo
        {
            FileName = fileName,
            Arguments = arguments,
            WorkingDirectory = workingDirectory,
            RedirectStandardOutput = true,
            RedirectStandardError = true,
            StandardOutputEncoding = Encoding.UTF8,
            StandardErrorEncoding = Encoding.UTF8,
            UseShellExecute = false,
            CreateNoWindow = true
        };

        if (environment is not null)
        {
            foreach (var item in environment)
            {
                process.StartInfo.Environment[item.Key] = item.Value;
            }
        }

        process.Start();
        var stdout = await process.StandardOutput.ReadToEndAsync();
        var stderr = await process.StandardError.ReadToEndAsync();
        await process.WaitForExitAsync();

        return new ProcessResult(process.ExitCode, stdout, stderr);
    }

    public static TempDirectory CreateTempDirectory() => new();

    public static string CreateSyntheticTemplateDocx(
        string directory,
        string fileName = "template.docx",
        bool includeReferenceParagraph = true,
        string titleHeading = "1 INTRODUCAO",
        string titleBody = "Texto exemplo do template.",
        string abstractHeading = "RESUMO",
        string abstractBody = "Conteudo exemplo do resumo.",
        string preservedRegionText = "REGIAO PRESERVADA DO TEMPLATE")
    {
        var docxPath = Path.Combine(directory, fileName);
        CreateDocument(
            docxPath,
            includeReferenceParagraph,
            titleHeading,
            titleBody,
            abstractHeading,
            abstractBody,
            preservedRegionText);
        return docxPath;
    }

    public static string CreateSyntheticSourceDocx(
        string directory,
        string fileName = "source.docx",
        bool includeReferenceParagraph = true,
        string titleHeading = "1 INTRODUCAO",
        string titleBody = "Texto fonte adaptado.",
        string abstractHeading = "RESUMO",
        string abstractBody = "Resumo de origem do documento.",
        string preservedRegionText = "REGIAO DE ORIGEM QUE DEVE SER MANTIDA")
    {
        var docxPath = Path.Combine(directory, fileName);
        CreateDocument(
            docxPath,
            includeReferenceParagraph,
            titleHeading,
            titleBody,
            abstractHeading,
            abstractBody,
            preservedRegionText);
        return docxPath;
    }

    public static string WriteCanonicalProfile(string directory, string templatePath, string profileFileName = "profile.canonical.json", string? templateHash = null)
    {
        var profilePath = Path.Combine(directory, profileFileName);
        var profile = new
        {
            template = new
            {
                name = "template-profile-sintetico",
                sourceFile = Path.GetFileName(templatePath),
                sha256 = templateHash ?? ComputeSha256(templatePath),
                profileVersion = 1
            },
            regions = new object[]
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
                    startBlockId = "p-0003",
                    endBlockId = "p-0004",
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
            }
        };

        File.WriteAllText(profilePath, JsonSerializer.Serialize(profile, new JsonSerializerOptions { WriteIndented = true }), Encoding.UTF8);
        return profilePath;
    }

    public static string WriteInvalidProfile(string directory, string templatePath, string profileFileName = "profile.invalid.json")
    {
        var profilePath = Path.Combine(directory, profileFileName);
        var profile = new
        {
            template = new
            {
                name = "template-profile-sintetico",
                sourceFile = Path.GetFileName(templatePath),
                sha256 = "0000000000000000000000000000000000000000000000000000000000000000",
                profileVersion = 1
            },
            regions = new object[]
            {
                new
                {
                    role = "title",
                    templateBlockId = "p-0002",
                    replaceWith = "source.title",
                    preserveFormatting = true
                }
            }
        };

        File.WriteAllText(profilePath, JsonSerializer.Serialize(profile, new JsonSerializerOptions { WriteIndented = true }), Encoding.UTF8);
        return profilePath;
    }

    public static string ComputeSha256(string filePath)
    {
        using var stream = File.OpenRead(filePath);
        var hash = SHA256.HashData(stream);
        return Convert.ToHexString(hash).ToLowerInvariant();
    }

    public static void MutateTemplateBodyForHashMismatch(string templatePath)
    {
        using var document = WordprocessingDocument.Open(templatePath, true);
        var body = document.MainDocumentPart!.Document.Body!;
        body.Append(new Paragraph(new Run(new Text("Paragrafo extra para alterar o hash."))));
        document.MainDocumentPart.Document.Save();
    }

    private static void CreateDocument(
        string docxPath,
        bool includeReferenceParagraph,
        string titleHeading,
        string titleBody,
        string abstractHeading,
        string abstractBody,
        string preservedRegionText)
    {
        using var document = WordprocessingDocument.Create(docxPath, WordprocessingDocumentType.Document);
        var mainPart = document.AddMainDocumentPart();
        mainPart.Document = new Document(new Body());

        var stylesPart = mainPart.AddNewPart<StyleDefinitionsPart>();
        stylesPart.Styles = CreateStyles();

        var body = mainPart.Document.Body!;
        body.Append(CreateParagraph(titleHeading, bold: true, allCaps: true, justification: JustificationValues.Center, spacingAfter: "120"));
        body.Append(CreateParagraph(titleBody, spacingAfter: "120"));
        body.Append(CreateParagraph(abstractHeading, bold: true, allCaps: true, spacingAfter: "120"));
        body.Append(CreateParagraph(abstractBody, spacingAfter: "120"));

        if (includeReferenceParagraph)
        {
            body.Append(CreateReferenceParagraph());
        }

        body.Append(CreateParagraph(preservedRegionText, spacingAfter: "120"));

        mainPart.Document.Save();
        stylesPart.Styles!.Save();
    }

    private static Styles CreateStyles()
    {
        var styles = new Styles();
        styles.Append(new Style(
            new StyleName { Val = "Normal" },
            new PrimaryStyle(),
            new StyleParagraphProperties(new SpacingBetweenLines { Before = "0", After = "0" })
        )
        {
            Type = StyleValues.Paragraph,
            StyleId = "Normal",
            Default = true
        });
        return styles;
    }

    private static Paragraph CreateParagraph(
        string text,
        bool bold = false,
        bool italic = false,
        bool allCaps = false,
        JustificationValues? justification = null,
        string? spacingAfter = null)
    {
        var paragraphProperties = new ParagraphProperties(new ParagraphStyleId { Val = "Normal" });
        if (justification is not null)
        {
            paragraphProperties.Append(new Justification { Val = justification.Value });
        }

        if (spacingAfter is not null)
        {
            paragraphProperties.Append(new SpacingBetweenLines { After = spacingAfter });
        }

        var runProperties = new RunProperties();
        if (bold)
        {
            runProperties.Append(new Bold());
        }

        if (italic)
        {
            runProperties.Append(new Italic());
        }

        if (allCaps)
        {
            runProperties.Append(new Caps());
        }

        var paragraph = new Paragraph(paragraphProperties);
        paragraph.Append(new Run(runProperties, new Text(text) { Space = SpaceProcessingModeValues.Preserve }));
        return paragraph;
    }

    private static Paragraph CreateReferenceParagraph()
    {
        var paragraph = new Paragraph(new ParagraphProperties(
            new ParagraphStyleId { Val = "Normal" },
            new SpacingBetweenLines { After = "120" },
            new Indentation { Hanging = "360" }));

        paragraph.Append(new Run(new RunProperties(new Caps()), new Text("SILVA, J. A. ") { Space = SpaceProcessingModeValues.Preserve }));
        paragraph.Append(new Run(new RunProperties(new Bold()), new Text("Titulo do artigo. ") { Space = SpaceProcessingModeValues.Preserve }));
        paragraph.Append(new Run(new RunProperties(new Italic()), new Text("Revista X") { Space = SpaceProcessingModeValues.Preserve }));
        paragraph.Append(new Run(new Text(", v. 10, n. 2, p. 1-10, 2024. ") { Space = SpaceProcessingModeValues.Preserve }));

        var hyperlinkRun = new Run(new RunProperties(new Color { Val = "0563C1" }, new Underline { Val = UnderlineValues.Single }));
        hyperlinkRun.Append(new Text("https://doi.org/10.1000/teste") { Space = SpaceProcessingModeValues.Preserve });
        paragraph.Append(hyperlinkRun);

        return paragraph;
    }

    public sealed record ProcessResult(int ExitCode, string StandardOutput, string StandardError)
    {
        public string CombinedOutput => StandardOutput + Environment.NewLine + StandardError;
    }

    public sealed class TempDirectory : IDisposable
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
