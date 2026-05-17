using System.Text;
using System.Text.Json;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Xunit;

namespace DocxOpenXmlTools.Tests;

public sealed class TemplateRealisticApplicationCliTests
{
    [Fact]
    public async Task Apply_template_skips_empty_source_candidates_before_title_and_abstract()
    {
        using var tempDir = TemplateCliTestSupport.CreateTempDirectory();
        var templatePath = CreateTemplateWithTitleAndAbstract(tempDir.Path, "template-realista-title-abstract.docx");
        var sourcePath = CreateSourceWithLeadingEmptyParagraphsAndWhitespaceAbstractBody(tempDir.Path, "source-realista-title-abstract.docx");
        var profilePath = WriteProfile(
            tempDir.Path,
            templatePath,
            new object[]
            {
                new
                {
                    role = "title",
                    templateBlockId = "p-0001",
                    replaceWith = "source.title",
                    preserveFormatting = true
                },
                new
                {
                    role = "abstract",
                    startBlockId = "p-0002",
                    endBlockId = "p-0003",
                    replaceWith = "source.abstract",
                    preserveFormattingFrom = "p-0002"
                }
            },
            "profile-title-abstract-realista.json");
        var outputPath = Path.Combine(tempDir.Path, "aplicado-title-abstract-realista.docx");
        var reportPath = Path.Combine(tempDir.Path, "aplicado-title-abstract-realista-report.md");

        var result = await TemplateCliTestSupport.RunToolsAsync(
            $"apply-template --template \"{templatePath}\" --source \"{sourcePath}\" --profile \"{profilePath}\" --out \"{outputPath}\" --report \"{reportPath}\"");

        Assert.Equal(0, result.ExitCode);
        Assert.True(File.Exists(outputPath), "O DOCX aplicado nao foi gerado.");
        Assert.True(File.Exists(reportPath), "O relatorio de aplicacao nao foi gerado.");

        using var document = WordprocessingDocument.Open(outputPath, false);
        var paragraphTexts = document.MainDocumentPart!.Document.Body!
            .Descendants<Paragraph>()
            .Select(paragraph => string.Concat(paragraph.Descendants<Text>().Select(text => text.Text)).Trim())
            .Where(text => !string.IsNullOrWhiteSpace(text))
            .ToArray();

        Assert.Contains(paragraphTexts, text => text.Contains("Título aplicado em fonte realista", StringComparison.OrdinalIgnoreCase));
        Assert.Contains(paragraphTexts, text => text.Contains("Resumo aplicado a partir de um candidato nao vazio", StringComparison.OrdinalIgnoreCase));

        var reportText = await File.ReadAllTextAsync(reportPath, Encoding.UTF8);
        Assert.DoesNotContain("No source content found for region `title`.", reportText, StringComparison.OrdinalIgnoreCase);
        Assert.DoesNotContain("No source content found for region `abstract`.", reportText, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public async Task Apply_template_resolves_reference_region_inside_nested_table_paragraph()
    {
        using var tempDir = TemplateCliTestSupport.CreateTempDirectory();
        var templatePath = CreateTemplateWithNestedReferenceRegion(tempDir.Path, "template-realista-references.docx");
        var sourcePath = CreateSourceWithReferenceParagraph(tempDir.Path, "source-realista-references.docx");
        var profilePath = WriteProfile(
            tempDir.Path,
            templatePath,
            new object[]
            {
                new
                {
                    role = "references",
                    templateBlockId = "p-0003",
                    replaceWith = "source.references",
                    referenceFormattingProfile = "refs-main"
                }
            },
            "profile-references-realista.json");
        var outputPath = Path.Combine(tempDir.Path, "aplicado-references-realista.docx");
        var reportPath = Path.Combine(tempDir.Path, "aplicado-references-realista-report.md");

        var result = await TemplateCliTestSupport.RunToolsAsync(
            $"apply-template --template \"{templatePath}\" --source \"{sourcePath}\" --profile \"{profilePath}\" --out \"{outputPath}\" --report \"{reportPath}\"");

        Assert.Equal(0, result.ExitCode);
        Assert.True(File.Exists(outputPath), "O DOCX aplicado nao foi gerado.");
        Assert.True(File.Exists(reportPath), "O relatorio de aplicacao nao foi gerado.");

        using var document = WordprocessingDocument.Open(outputPath, false);
        var referenceParagraphTexts = document.MainDocumentPart!.Document.Body!
            .Descendants<Paragraph>()
            .Select(paragraph => string.Concat(paragraph.Descendants<Text>().Select(text => text.Text)).Trim())
            .Where(text => !string.IsNullOrWhiteSpace(text))
            .ToArray();

        Assert.Contains(referenceParagraphTexts, text =>
            text.Contains("SILVA, J. A.", StringComparison.OrdinalIgnoreCase) &&
            text.Contains("Titulo do artigo.", StringComparison.OrdinalIgnoreCase) &&
            text.Contains("Revista X", StringComparison.OrdinalIgnoreCase));

        var reportText = await File.ReadAllTextAsync(reportPath, Encoding.UTF8);
        Assert.DoesNotContain("No template block found for region `references`.", reportText, StringComparison.OrdinalIgnoreCase);
    }

    private static string CreateTemplateWithTitleAndAbstract(string directory, string fileName)
    {
        return CreateDocument(
            directory,
            fileName,
            [
                CreateEmptyParagraph(),
                CreateEmptyParagraph(),
                CreateParagraph("Título do template realista", bold: true, allCaps: true, justification: JustificationValues.Center, spacingAfter: "120"),
                CreateParagraph("Texto-base do título do template.", spacingAfter: "120"),
                CreateParagraph("RESUMO", bold: true, allCaps: true, spacingAfter: "120"),
                CreateParagraph("Texto-base do resumo do template.", spacingAfter: "120")
            ]);
    }

    private static string CreateSourceWithLeadingEmptyParagraphsAndWhitespaceAbstractBody(string directory, string fileName)
    {
        return CreateDocument(
            directory,
            fileName,
            [
                CreateEmptyParagraph(),
                CreateEmptyParagraph(),
                CreateParagraph("Título aplicado em fonte realista", bold: true, allCaps: true, justification: JustificationValues.Center, spacingAfter: "120"),
                CreateParagraph("RESUMO", bold: true, allCaps: true, spacingAfter: "120"),
                CreateParagraph("   ", spacingAfter: "120"),
                CreateParagraph("Resumo aplicado a partir de um candidato nao vazio.", spacingAfter: "120")
            ]);
    }

    private static string CreateTemplateWithNestedReferenceRegion(string directory, string fileName)
    {
        return CreateDocument(
            directory,
            fileName,
            [
                CreateParagraph("Titulo do template de referencias", bold: true, allCaps: true, justification: JustificationValues.Center, spacingAfter: "120"),
                CreateParagraph("RESUMO", bold: true, allCaps: true, spacingAfter: "120"),
                CreateTableWithSingleReferenceParagraph()
            ]);
    }

    private static string CreateSourceWithReferenceParagraph(string directory, string fileName)
    {
        return CreateDocument(
            directory,
            fileName,
            [
                CreateParagraph("Titulo da fonte de referencias", bold: true, allCaps: true, justification: JustificationValues.Center, spacingAfter: "120"),
                CreateParagraph("RESUMO", bold: true, allCaps: true, spacingAfter: "120"),
                CreateReferenceParagraph()
            ]);
    }

    private static Table CreateTableWithSingleReferenceParagraph()
    {
        return new Table(
            new TableProperties(
                new TableWidth { Width = "5000", Type = TableWidthUnitValues.Pct },
                new TableLook { Val = "04A0" }),
            new TableRow(
                new TableCell(
                    new TableCellProperties(new TableCellWidth { Width = "5000", Type = TableWidthUnitValues.Pct }),
                    CreateReferenceParagraph())));
    }

    private static string CreateDocument(
        string directory,
        string fileName,
        IEnumerable<OpenXmlElement> bodyElements)
    {
        var docxPath = Path.Combine(directory, fileName);
        using var document = WordprocessingDocument.Create(docxPath, WordprocessingDocumentType.Document);
        var mainPart = document.AddMainDocumentPart();
        mainPart.Document = new Document(new Body());

        var stylesPart = mainPart.AddNewPart<StyleDefinitionsPart>();
        stylesPart.Styles = CreateStyles();

        var body = mainPart.Document.Body!;
        foreach (var element in bodyElements)
        {
            body.Append((OpenXmlElement)element.CloneNode(true));
        }

        mainPart.Document.Save();
        stylesPart.Styles!.Save();
        return docxPath;
    }

    private static Styles CreateStyles()
    {
        var styles = new Styles();
        styles.Append(new Style(
            new StyleName { Val = "Normal" },
            new PrimaryStyle(),
            new StyleParagraphProperties(new SpacingBetweenLines { Before = "0", After = "0" }))
        {
            Type = StyleValues.Paragraph,
            StyleId = "Normal",
            Default = true
        });
        return styles;
    }

    private static Paragraph CreateEmptyParagraph()
    {
        return new Paragraph(new ParagraphProperties(new ParagraphStyleId { Val = "Normal" }));
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

    private static string WriteProfile(string directory, string templatePath, object[] regions, string profileFileName)
    {
        var profilePath = Path.Combine(directory, profileFileName);
        var profile = new
        {
            template = new
            {
                name = "template-profile-realista",
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
