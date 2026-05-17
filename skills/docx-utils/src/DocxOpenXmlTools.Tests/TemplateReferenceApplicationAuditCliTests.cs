using System.Text;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Xunit;

namespace DocxOpenXmlTools.Tests;

public sealed class TemplateReferenceApplicationAuditCliTests
{
    [Fact]
    public async Task Apply_template_and_audit_template_application_accept_references_heading_plus_first_bibliographic_paragraph()
    {
        using var tempDir = TemplateCliTestSupport.CreateTempDirectory();
        var templatePath = TemplateCliTestSupport.CreateSyntheticTemplateDocx(
            tempDir.Path,
            fileName: "template-references-audit.docx",
            preservedRegionText: "BLOCO PRESERVADO DO TEMPLATE");
        var sourcePath = TemplateCliTestSupport.CreateSyntheticSourceDocx(
            tempDir.Path,
            fileName: "source-references-audit.docx",
            preservedRegionText: "BLOCO PRESERVADO DA ORIGEM");
        var outputPath = Path.Combine(tempDir.Path, "aplicado-references-audit.docx");
        var reportPath = Path.Combine(tempDir.Path, "audit-references-audit.md");

        var profilePath = TemplateCliTestSupport.WriteCanonicalProfile(tempDir.Path, templatePath);

        InsertParagraphBeforeOrdinal(
            sourcePath,
            5,
            CreateParagraph("Texto instrucional da secao de referencias.", spacingAfter: "120"));
        RewriteParagraphText(sourcePath, 6, "REFERÊNCIAS", bold: true, allCaps: true, spacingAfter: "120");
        InsertParagraphAfterOrdinal(sourcePath, 6, CreateWhitespaceParagraph());
        InsertParagraphAfterOrdinal(
            sourcePath,
            7,
            CreateReferenceParagraph(
                "AGÊNCIA NACIONAL DE ENERGIA ELÉTRICA.",
                "Banco de dados de energia elétrica."));

        var applyResult = await TemplateCliTestSupport.RunToolsAsync(
            $"apply-template --template \"{templatePath}\" --source \"{sourcePath}\" --profile \"{profilePath}\" --out \"{outputPath}\"");

        Assert.True(applyResult.ExitCode == 0, applyResult.CombinedOutput);
        Assert.True(File.Exists(outputPath), "O DOCX aplicado nao foi gerado.");

        using (var document = WordprocessingDocument.Open(outputPath, false))
        {
            var paragraphTexts = document.MainDocumentPart!.Document.Body!
                .Elements<Paragraph>()
                .Select(paragraph => string.Concat(paragraph.Descendants<Text>().Select(text => text.Text)))
                .ToArray();

            var referencesHeadingIndex = Array.FindIndex(
                paragraphTexts,
                text => string.Equals(text.Trim(), "REFERÊNCIAS", StringComparison.OrdinalIgnoreCase));

            Assert.True(referencesHeadingIndex >= 0, "Nao foi possivel localizar a secao REFERÊNCIAS no DOCX aplicado.");

            var firstBibliographicParagraphAfterHeading = paragraphTexts
                .Skip(referencesHeadingIndex + 1)
                .FirstOrDefault(text => !string.IsNullOrWhiteSpace(text));

            Assert.NotNull(firstBibliographicParagraphAfterHeading);
            Assert.Contains("AGÊNCIA NACIONAL DE ENERGIA ELÉTRICA", firstBibliographicParagraphAfterHeading, StringComparison.OrdinalIgnoreCase);
            Assert.DoesNotContain("Texto instrucional da secao de referencias.", firstBibliographicParagraphAfterHeading, StringComparison.OrdinalIgnoreCase);
        }

        var auditResult = await TemplateCliTestSupport.RunToolsAsync(
            $"audit-template-application \"{outputPath}\" --profile \"{profilePath}\" --report \"{reportPath}\"");

        Assert.True(auditResult.ExitCode == 0, auditResult.CombinedOutput);
        Assert.True(File.Exists(reportPath), "O relatorio de auditoria deveria ser gerado.");

        var reportText = await File.ReadAllTextAsync(reportPath, Encoding.UTF8);
        Assert.Contains("Pendencias: 0", reportText, StringComparison.OrdinalIgnoreCase);
        Assert.DoesNotContain("p-0623", reportText, StringComparison.OrdinalIgnoreCase);
        Assert.DoesNotContain("p-0624", reportText, StringComparison.OrdinalIgnoreCase);
    }

    private static void RewriteParagraphText(
        string docxPath,
        int ordinal,
        string text,
        bool bold = false,
        bool allCaps = false,
        string? spacingAfter = null)
    {
        using var document = WordprocessingDocument.Open(docxPath, true);
        var body = document.MainDocumentPart!.Document.Body!;
        var paragraph = body.Elements<Paragraph>().ElementAt(ordinal - 1);
        paragraph.RemoveAllChildren<Run>();

        if (paragraph.ParagraphProperties is null)
        {
            paragraph.PrependChild(new ParagraphProperties(new ParagraphStyleId { Val = "Normal" }));
        }

        if (spacingAfter is not null && paragraph.ParagraphProperties is not null)
        {
            var spacing = paragraph.ParagraphProperties.GetFirstChild<SpacingBetweenLines>();
            if (spacing is null)
            {
                paragraph.ParagraphProperties.Append(new SpacingBetweenLines { After = spacingAfter });
            }
            else
            {
                spacing.After = spacingAfter;
            }
        }

        var runProperties = new RunProperties();
        if (bold)
        {
            runProperties.Append(new Bold());
        }

        if (allCaps)
        {
            runProperties.Append(new Caps());
        }

        paragraph.Append(new Run(runProperties, new Text(text) { Space = SpaceProcessingModeValues.Preserve }));
        document.MainDocumentPart.Document.Save();
    }

    private static void InsertParagraphBeforeOrdinal(string docxPath, int ordinal, Paragraph paragraph)
    {
        using var document = WordprocessingDocument.Open(docxPath, true);
        var body = document.MainDocumentPart!.Document.Body!;
        var target = body.Elements<Paragraph>().ElementAt(ordinal - 1);
        target.InsertBeforeSelf(paragraph);
        document.MainDocumentPart.Document.Save();
    }

    private static void InsertParagraphAfterOrdinal(string docxPath, int ordinal, Paragraph paragraphToInsert)
    {
        using var document = WordprocessingDocument.Open(docxPath, true);
        var body = document.MainDocumentPart!.Document.Body!;
        var target = body.Elements<Paragraph>().ElementAt(ordinal - 1);
        target.InsertAfterSelf(paragraphToInsert);
        document.MainDocumentPart.Document.Save();
    }

    private static Paragraph CreateParagraph(
        string text,
        bool bold = false,
        bool allCaps = false,
        string? spacingAfter = null)
    {
        var paragraphProperties = new ParagraphProperties(new ParagraphStyleId { Val = "Normal" });
        if (spacingAfter is not null)
        {
            paragraphProperties.Append(new SpacingBetweenLines { After = spacingAfter });
        }

        var runProperties = new RunProperties();
        if (bold)
        {
            runProperties.Append(new Bold());
        }

        if (allCaps)
        {
            runProperties.Append(new Caps());
        }

        var paragraph = new Paragraph(paragraphProperties);
        paragraph.Append(new Run(runProperties, new Text(text) { Space = SpaceProcessingModeValues.Preserve }));
        return paragraph;
    }

    private static Paragraph CreateWhitespaceParagraph()
    {
        var paragraph = new Paragraph(new ParagraphProperties(
            new ParagraphStyleId { Val = "Normal" },
            new SpacingBetweenLines { After = "120" }));
        paragraph.Append(new Run(new Text(" ") { Space = SpaceProcessingModeValues.Preserve }));
        return paragraph;
    }

    private static Paragraph CreateReferenceParagraph(string authorText, string titleText)
    {
        var paragraph = new Paragraph(new ParagraphProperties(
            new ParagraphStyleId { Val = "Normal" },
            new SpacingBetweenLines { After = "120" },
            new Indentation { Hanging = "360" }));

        paragraph.Append(new Run(new RunProperties(new Caps()), new Text(authorText + " ") { Space = SpaceProcessingModeValues.Preserve }));
        paragraph.Append(new Run(new RunProperties(new Bold()), new Text(titleText + " ") { Space = SpaceProcessingModeValues.Preserve }));
        paragraph.Append(new Run(new RunProperties(new Italic()), new Text("Revista X") { Space = SpaceProcessingModeValues.Preserve }));
        paragraph.Append(new Run(new Text(", v. 10, n. 2, p. 1-10, 2024. ") { Space = SpaceProcessingModeValues.Preserve }));

        var hyperlinkRun = new Run(new RunProperties(new Color { Val = "0563C1" }, new Underline { Val = UnderlineValues.Single }));
        hyperlinkRun.Append(new Text("https://doi.org/10.1000/teste") { Space = SpaceProcessingModeValues.Preserve });
        paragraph.Append(hyperlinkRun);

        return paragraph;
    }
}
