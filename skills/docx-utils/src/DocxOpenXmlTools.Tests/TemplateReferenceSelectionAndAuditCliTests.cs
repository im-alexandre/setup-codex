using System.Text;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Xunit;

namespace DocxOpenXmlTools.Tests;

public sealed class TemplateReferenceSelectionAndAuditCliTests
{
    [Fact]
    public async Task Apply_template_deve_usar_o_primeiro_paragrafo_bibliografico_apos_referencias()
    {
        using var tempDir = TemplateCliTestSupport.CreateTempDirectory();
        var templatePath = TemplateCliTestSupport.CreateSyntheticTemplateDocx(
            tempDir.Path,
            fileName: "template-referencias.docx",
            preservedRegionText: "BLOCO PRESERVADO DO TEMPLATE");
        var sourcePath = TemplateCliTestSupport.CreateSyntheticSourceDocx(
            tempDir.Path,
            fileName: "source-referencias.docx",
            preservedRegionText: "BLOCO PRESERVADO DA ORIGEM");
        var profilePath = TemplateCliTestSupport.WriteCanonicalProfile(tempDir.Path, templatePath);
        var outputPath = Path.Combine(tempDir.Path, "aplicado-referencias.docx");

        ReplaceParagraphText(sourcePath, 5, "REFERÊNCIAS", bold: true, allCaps: true, spacingAfter: "120");
        InsertParagraphBeforeOrdinal(sourcePath, 5, CreateParagraph("Trecho referente a referencias bibliograficas, mas ainda nao e a secao final.", spacingAfter: "120"));
        InsertParagraphAfterOrdinal(sourcePath, 6, CreateWhitespaceParagraph());
        InsertParagraphAfterOrdinal(sourcePath, 7, CreateReferenceParagraph("AGÊNCIA NACIONAL DE ENERGIA ELÉTRICA.", "Banco de dados de energia elétrica."));
        InsertParagraphAfterOrdinal(sourcePath, 8, CreateReferenceParagraph("SILVA, J. A.", "Titulo do artigo."));

        var applyResult = await TemplateCliTestSupport.RunToolsAsync(
            $"apply-template --template \"{templatePath}\" --source \"{sourcePath}\" --profile \"{profilePath}\" --out \"{outputPath}\"");

        Assert.Equal(0, applyResult.ExitCode);
        Assert.True(File.Exists(outputPath), "O DOCX aplicado nao foi gerado.");

        var paragraphTexts = GetParagraphTexts(outputPath);

        var referencesHeadingIndex = Array.FindIndex(
            paragraphTexts,
            text => string.Equals(text.Trim(), "REFERÊNCIAS", StringComparison.OrdinalIgnoreCase));

        Assert.True(referencesHeadingIndex >= 0, "Nao foi possivel localizar a secao REFERÊNCIAS no DOCX aplicado.");
        Assert.Contains(paragraphTexts.Take(referencesHeadingIndex), text => text.Contains("trecho referente", StringComparison.OrdinalIgnoreCase));

        var firstBibliographicParagraphAfterHeading = paragraphTexts
            .Skip(referencesHeadingIndex + 1)
            .FirstOrDefault(text => !string.IsNullOrWhiteSpace(text));

        Assert.NotNull(firstBibliographicParagraphAfterHeading);
        Assert.Contains("AGÊNCIA NACIONAL DE ENERGIA ELÉTRICA", firstBibliographicParagraphAfterHeading, StringComparison.OrdinalIgnoreCase);
        Assert.DoesNotContain("trecho referente", firstBibliographicParagraphAfterHeading, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public async Task Audit_template_application_nao_deve_reportar_whitespace_preservado_mas_deve_continuar_detectando_texto_real()
    {
        using var tempDir = TemplateCliTestSupport.CreateTempDirectory();
        var templatePath = TemplateCliTestSupport.CreateSyntheticTemplateDocx(
            tempDir.Path,
            fileName: "template-auditoria.docx",
            preservedRegionText: "BLOCO PRESERVADO DO TEMPLATE");
        var sourcePath = TemplateCliTestSupport.CreateSyntheticSourceDocx(
            tempDir.Path,
            fileName: "source-auditoria.docx",
            preservedRegionText: "BLOCO PRESERVADO DA ORIGEM");
        var outputPath = Path.Combine(tempDir.Path, "aplicado-auditoria.docx");
        var reportPath = Path.Combine(tempDir.Path, "auditoria.md");

        ReplaceParagraphText(templatePath, 4, string.Empty);
        ReplaceParagraphText(sourcePath, 4, string.Empty);
        ReplaceParagraphText(sourcePath, 5, "AGÊNCIA NACIONAL DE ENERGIA ELÉTRICA. Banco de dados de energia elétrica.", spacingAfter: "120");
        var profilePath = TemplateCliTestSupport.WriteCanonicalProfile(tempDir.Path, templatePath);

        var applyResult = await TemplateCliTestSupport.RunToolsAsync(
            $"apply-template --template \"{templatePath}\" --source \"{sourcePath}\" --profile \"{profilePath}\" --out \"{outputPath}\"");

        Assert.Equal(0, applyResult.ExitCode);
        Assert.True(File.Exists(outputPath), "O DOCX aplicado nao foi gerado.");

        AppendParagraph(outputPath, "REGIAO PRESERVADA DO TEMPLATE");

        var auditResult = await TemplateCliTestSupport.RunToolsAsync(
            $"audit-template-application \"{outputPath}\" --profile \"{profilePath}\" --report \"{reportPath}\"");

        Assert.NotEqual(0, auditResult.ExitCode);
        Assert.True(File.Exists(reportPath), "O relatorio de auditoria nao foi gerado.");

        var reportText = await File.ReadAllTextAsync(reportPath, Encoding.UTF8);
        Assert.Contains("REGIAO PRESERVADA DO TEMPLATE", reportText, StringComparison.OrdinalIgnoreCase);
        Assert.DoesNotContain("p-0004", reportText, StringComparison.OrdinalIgnoreCase);
    }

    private static void InsertParagraphBeforeOrdinal(string docxPath, int ordinal, Paragraph paragraph)
    {
        using var document = WordprocessingDocument.Open(docxPath, true);
        var body = document.MainDocumentPart!.Document.Body!;
        var target = body.Elements<Paragraph>().ElementAt(ordinal - 1);
        target.InsertBeforeSelf(paragraph);
        document.MainDocumentPart.Document.Save();
    }

    private static void InsertParagraphAfterOrdinal(string docxPath, int ordinal, Paragraph paragraph)
    {
        using var document = WordprocessingDocument.Open(docxPath, true);
        var body = document.MainDocumentPart!.Document.Body!;
        var target = body.Elements<Paragraph>().ElementAt(ordinal - 1);
        target.InsertAfterSelf(paragraph);
        document.MainDocumentPart.Document.Save();
    }

    private static void ReplaceParagraphText(
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

        var props = new RunProperties();
        if (bold)
        {
            props.Append(new Bold());
        }

        if (allCaps)
        {
            props.Append(new Caps());
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

        paragraph.Append(new Run(props, new Text(text) { Space = SpaceProcessingModeValues.Preserve }));
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

    private static void AppendParagraph(string docxPath, string text)
    {
        using var document = WordprocessingDocument.Open(docxPath, true);
        var body = document.MainDocumentPart!.Document.Body!;
        body.Append(CreateParagraph(text, spacingAfter: "120"));
        document.MainDocumentPart.Document.Save();
    }

    private static string[] GetParagraphTexts(string docxPath)
    {
        using var document = WordprocessingDocument.Open(docxPath, false);
        return document.MainDocumentPart!.Document.Body!
            .Elements<Paragraph>()
            .Select(paragraph => string.Concat(paragraph.Descendants<Text>().Select(text => text.Text)))
            .ToArray();
    }
}
