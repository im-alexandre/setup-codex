using System.Globalization;
using System.Security.Cryptography;
using System.Text;
using System.Text.RegularExpressions;
using DocxOpenXmlTools.TemplateProfiles;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace DocxOpenXmlTools.TemplateProfiles.Inspection;

internal static class TemplateDocumentInspector
{
    private static readonly Regex ManualNumberingPrefix = new(
        @"^\s*(\d+(?:\.\d+)*)[.)]?\s+",
        RegexOptions.Compiled | RegexOptions.CultureInvariant);

    public static TemplateInspectionDocument Inspect(string templatePath)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(templatePath);

        var resolvedPath = Path.GetFullPath(templatePath);
        if (!File.Exists(resolvedPath))
        {
            throw new FileNotFoundException("Template DOCX nao encontrado.", resolvedPath);
        }

        using var document = WordprocessingDocument.Open(resolvedPath, false);
        var candidates = new List<TemplateInspectionCandidate>();

        foreach (var partDocument in EnumerateDocumentParts(document))
        {
            foreach (var paragraph in partDocument.Root.Descendants<Paragraph>())
            {
                candidates.Add(BuildCandidate(paragraph, partDocument.PartName, candidates.Count + 1));
            }
        }

        return new TemplateInspectionDocument
        {
            Template = new TemplateInspectionTemplate
            {
                Name = Path.GetFileNameWithoutExtension(resolvedPath),
                SourceFile = Path.GetFileName(resolvedPath),
                Sha256 = ComputeSha256(resolvedPath),
                ProfileVersion = 1
            },
            Candidates = candidates
        };
    }

    private static TemplateInspectionCandidate BuildCandidate(Paragraph paragraph, string partName, int ordinal)
    {
        var text = ExtractParagraphText(paragraph);
        var paragraphProperties = paragraph.ParagraphProperties;
        var runs = paragraph.Descendants<Run>().Select(BuildRun).ToArray();
        var styleId = paragraphProperties?.ParagraphStyleId?.Val?.Value ?? "Normal";

        return new TemplateInspectionCandidate
        {
            Id = $"p-{ordinal:0000}",
            Text = text,
            Ordinal = ordinal,
            Location = new TemplateInspectionLocation
            {
                Part = ResolveLocationPart(partName, paragraph),
                Ordinal = ordinal
            },
            ParagraphFormat = new TemplateInspectionParagraphFormat
            {
                StyleId = styleId,
                Alignment = ResolveAlignment(paragraphProperties?.Justification?.Val?.Value),
                SpacingBefore = TryParseInt(paragraphProperties?.SpacingBetweenLines?.Before?.Value),
                SpacingAfter = TryParseInt(paragraphProperties?.SpacingBetweenLines?.After?.Value),
                HangingIndent = TryParseInt(paragraphProperties?.Indentation?.Hanging?.Value)
            },
            Runs = runs,
            StructuralHints = new TemplateInspectionStructuralHints
            {
                ManualNumbering = TryGetManualNumbering(text),
                ShortHighlightedParagraph = IsShortHighlightedParagraph(text, runs),
                LooksLikeReference = LooksLikeReference(text, paragraphProperties, runs)
            }
        };
    }

    private static TemplateInspectionRun BuildRun(Run run)
    {
        var runProperties = run.RunProperties;
        var fontSize = TryParseInt(runProperties?.FontSize?.Val?.Value);

        return new TemplateInspectionRun
        {
            Text = ExtractRunText(run),
            Bold = runProperties?.Bold is not null,
            Italic = runProperties?.Italic is not null,
            AllCaps = runProperties?.Caps is not null || runProperties?.SmallCaps is not null,
            FontSize = fontSize is null ? null : NormalizeFontSize(fontSize.Value)
        };
    }

    private static IEnumerable<(OpenXmlElement Root, string PartName)> EnumerateDocumentParts(WordprocessingDocument document)
    {
        var mainPart = document.MainDocumentPart;
        if (mainPart?.Document?.Body is not null)
        {
            yield return (mainPart.Document.Body, "body");
        }

        foreach (var headerPart in mainPart?.HeaderParts ?? [])
        {
            if (headerPart.Header is not null)
            {
                yield return (headerPart.Header, "header");
            }
        }

        foreach (var footerPart in mainPart?.FooterParts ?? [])
        {
            if (footerPart.Footer is not null)
            {
                yield return (footerPart.Footer, "footer");
            }
        }

        if (mainPart?.FootnotesPart?.Footnotes is not null)
        {
            yield return (mainPart.FootnotesPart.Footnotes, "footnote");
        }

        if (mainPart?.EndnotesPart?.Endnotes is not null)
        {
            yield return (mainPart.EndnotesPart.Endnotes, "endnote");
        }
    }

    private static string ExtractParagraphText(OpenXmlElement paragraph)
    {
        var builder = new StringBuilder();
        foreach (var run in paragraph.Descendants<Run>())
        {
            builder.Append(ExtractRunText(run));
        }

        return builder.ToString();
    }

    private static string ExtractRunText(OpenXmlElement run)
    {
        var builder = new StringBuilder();
        foreach (var text in run.Descendants<Text>())
        {
            builder.Append(text.Text);
        }

        return builder.ToString();
    }

    private static string ResolveLocationPart(string partName, Paragraph paragraph)
    {
        if (string.Equals(partName, "body", StringComparison.OrdinalIgnoreCase) && paragraph.Ancestors<TableCell>().Any())
        {
            return "table-cell";
        }

        return partName.ToLowerInvariant();
    }

    private static string ResolveAlignment(JustificationValues? justification)
    {
        var value = justification?.ToString();
        return value switch
        {
            "center" => "center",
            "right" => "right",
            "both" => "justify",
            "distribute" => "distributed",
            "left" => "left",
            _ => "left"
        };
    }

    private static bool IsShortHighlightedParagraph(string text, IReadOnlyList<TemplateInspectionRun> runs)
    {
        return text.Length <= 60 && (TryGetManualNumbering(text) is not null || runs.Any(run => run.Bold || run.Italic || run.AllCaps));
    }

    private static bool LooksLikeReference(string text, ParagraphProperties? paragraphProperties, IReadOnlyList<TemplateInspectionRun> runs)
    {
        if (paragraphProperties?.Indentation?.Hanging is not null)
        {
            return true;
        }

        if (text.Contains("doi", StringComparison.OrdinalIgnoreCase) ||
            text.Contains("http", StringComparison.OrdinalIgnoreCase) ||
            text.Contains("www.", StringComparison.OrdinalIgnoreCase))
        {
            return true;
        }

        if (text.Count(char.IsDigit) >= 4 && runs.Any(run => run.Italic || run.Bold))
        {
            return true;
        }

        return false;
    }

    private static int NormalizeFontSize(int halfPoints)
    {
        return (int)Math.Round(halfPoints / 2.0, MidpointRounding.AwayFromZero);
    }

    private static string? TryGetManualNumbering(string text)
    {
        var match = ManualNumberingPrefix.Match(text);
        return match.Success ? match.Groups[1].Value : null;
    }

    private static int? TryParseInt(string? value)
    {
        if (int.TryParse(value, NumberStyles.Integer, CultureInfo.InvariantCulture, out var parsed))
        {
            return parsed;
        }

        return null;
    }

    private static string ComputeSha256(string filePath)
    {
        using var stream = File.OpenRead(filePath);
        var hash = SHA256.HashData(stream);
        return Convert.ToHexString(hash).ToLowerInvariant();
    }
}
