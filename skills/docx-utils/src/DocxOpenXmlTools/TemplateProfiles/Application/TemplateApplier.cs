using System.Text;
using DocumentFormat.OpenXml;
using DocxOpenXmlTools.TemplateProfiles.Inspection;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace DocxOpenXmlTools.TemplateProfiles.Application;

internal static class TemplateApplier
{
    public static TemplateApplicationResult Apply(
        string templatePath,
        string sourcePath,
        string profilePath,
        string outputPath,
        string? reportPath)
    {
        var errors = ValidateInputs(templatePath, sourcePath, profilePath, outputPath);
        if (errors.Count > 0)
        {
            return new TemplateApplicationResult(false, 4, new TemplateApplicationReport
            {
                Template = templatePath,
                Source = sourcePath,
                Profile = profilePath,
                Output = outputPath,
                PendingIssues = errors
            }, errors);
        }

        var validation = TemplateProfileValidator.Validate(profilePath);
        if (validation.Profile is null)
        {
            return new TemplateApplicationResult(false, 6, new TemplateApplicationReport
            {
                Template = templatePath,
                Source = sourcePath,
                Profile = profilePath,
                Output = outputPath,
                PendingIssues = validation.Errors
            }, validation.Errors);
        }

        var blockingValidationErrors = validation.Errors
            .Where(error => !IsMissingRequiredRegionError(error))
            .ToArray();
        if (blockingValidationErrors.Length > 0)
        {
            return new TemplateApplicationResult(false, 6, new TemplateApplicationReport
            {
                Template = templatePath,
                Source = sourcePath,
                Profile = profilePath,
                Output = outputPath,
                PendingIssues = blockingValidationErrors
            }, blockingValidationErrors);
        }

        var resolvedTemplatePath = Path.GetFullPath(templatePath);
        if (!string.Equals(Path.GetFullPath(validation.ResolvedTemplatePath), resolvedTemplatePath, StringComparison.OrdinalIgnoreCase))
        {
            return new TemplateApplicationResult(false, 6, new TemplateApplicationReport
            {
                Template = templatePath,
                Source = sourcePath,
                Profile = profilePath,
                Output = outputPath,
                PendingIssues = [$"Template path differs from profile sourceFile: {validation.ResolvedTemplatePath}"]
            }, [$"Template path differs from profile sourceFile: {validation.ResolvedTemplatePath}"]);
        }

        var normalizedOutput = Path.GetFullPath(outputPath);
        Directory.CreateDirectory(Path.GetDirectoryName(normalizedOutput) ?? ".");
        File.Copy(resolvedTemplatePath, normalizedOutput, overwrite: true);

        var sourceInspection = TemplateDocumentInspector.Inspect(sourcePath);
        var application = ApplyRegions(normalizedOutput, sourcePath, sourceInspection, validation.Profile);

        var report = new TemplateApplicationReport
        {
            Template = resolvedTemplatePath,
            Source = Path.GetFullPath(sourcePath),
            Profile = Path.GetFullPath(profilePath),
            Output = normalizedOutput,
            AppliedRegions = application.AppliedRegions,
            PreservedRegions = application.PreservedRegions,
            PendingIssues = application.PendingIssues
        };

        if (!string.IsNullOrWhiteSpace(reportPath))
        {
            var normalizedReportPath = Path.GetFullPath(reportPath);
            Directory.CreateDirectory(Path.GetDirectoryName(normalizedReportPath) ?? ".");
            File.WriteAllText(normalizedReportPath, TemplateApplicationReportWriter.BuildMarkdown(report), Encoding.UTF8);
        }

        var success = application.PendingIssues.Count == 0;
        return new TemplateApplicationResult(success, success ? 0 : 6, report, application.PendingIssues);
    }

    private static TemplateApplicationSummary ApplyRegions(
        string outputPath,
        string sourcePath,
        TemplateInspectionDocument sourceInspection,
        TemplateProfileDocument profile)
    {
        var applied = new List<string>();
        var preserved = new List<string>();
        var pending = new List<string>();

        using var document = WordprocessingDocument.Open(outputPath, true);
        using var sourceDocument = WordprocessingDocument.Open(sourcePath, false);
        var body = document.MainDocumentPart?.Document?.Body;
        var sourceBody = sourceDocument.MainDocumentPart?.Document?.Body;
        if (body is null || sourceBody is null)
        {
            return new TemplateApplicationSummary([], [], ["Output document has no body."]);
        }

        var paragraphs = body.Descendants<Paragraph>().ToList();
        foreach (var region in profile.Regions.OrderByDescending(GetRegionStartOrdinal))
        {
            var targetSpan = ResolveTargetParagraphSpan(region, paragraphs, out var targetError);
            if (targetSpan is null)
            {
                pending.Add(targetError ?? $"No template block found for region `{region.Role}`.");
                continue;
            }

            if (IsReferenceRegion(region))
            {
                var referenceSection = ResolveReferenceSection(sourceInspection);
                var referenceParagraph = referenceSection.Reference is null
                    ? null
                    : GetParagraphByOrdinal(sourceBody, referenceSection.Reference.Ordinal);
                if (referenceParagraph is null)
                {
                    pending.Add($"No source content found for region `{region.Role}`.");
                    continue;
                }

                preserved.AddRange(
                    targetSpan
                        .Skip(1)
                        .Select(GetParagraphText)
                        .Where(IsPreservedRegionText));

                if (referenceSection.Heading is not null && targetSpan.Count > 1)
                {
                    var headingParagraph = GetParagraphByOrdinal(sourceBody, referenceSection.Heading.Ordinal);
                    if (headingParagraph is null)
                    {
                        pending.Add($"No source content found for region `{region.Role}`.");
                        continue;
                    }

                    if (referenceSection.LeadIn is not null)
                    {
                        var leadInParagraph = GetParagraphByOrdinal(sourceBody, referenceSection.LeadIn.Ordinal);
                        if (leadInParagraph is null)
                        {
                            pending.Add($"No source content found for region `{region.Role}`.");
                            continue;
                        }

                        ReplaceParagraphWithClone(targetSpan[0], leadInParagraph);
                        ReplaceParagraphWithClone(targetSpan[1], headingParagraph);
                        InsertParagraphAfter(targetSpan[1], referenceParagraph);
                        ClearParagraphContent(targetSpan.Skip(2));
                    }
                    else
                    {
                        ReplaceParagraphWithClone(targetSpan[0], headingParagraph);
                        ReplaceParagraphWithClone(targetSpan[1], referenceParagraph);
                        ClearParagraphContent(targetSpan.Skip(2));
                    }
                }
                else
                {
                    ReplaceParagraphWithClone(targetSpan[0], referenceParagraph);
                    ClearParagraphContent(targetSpan.Skip(1));
                }
                applied.Add(region.Role);
                continue;
            }

            var replacement = ResolveReplacementText(region, sourceInspection);
            if (string.IsNullOrWhiteSpace(replacement))
            {
                pending.Add($"No source content found for region `{region.Role}`.");
                continue;
            }

            ReplaceParagraphTextPreservingFormat(targetSpan[0], replacement);
            ClearParagraphContent(targetSpan.Skip(1));
            applied.Add(region.Role);
        }

        document.MainDocumentPart!.Document.Save();

        preserved.AddRange(body.Elements<Paragraph>()
            .Select(p => string.Concat(p.Descendants<Text>().Select(t => t.Text)).Trim())
            .Where(text => !string.IsNullOrWhiteSpace(text) && !applied.Any(role => text.Contains(role, StringComparison.OrdinalIgnoreCase)))
            .Where(IsPreservedRegionText));

        return new TemplateApplicationSummary(
            applied,
            preserved.Distinct(StringComparer.OrdinalIgnoreCase).ToArray(),
            pending);
    }

    private static bool IsReferenceRegion(TemplateProfileRegion region)
    {
        return string.Equals(region.Role, "references", StringComparison.OrdinalIgnoreCase) ||
               !string.IsNullOrWhiteSpace(region.ReferenceFormattingProfile);
    }

    private static int GetRegionStartOrdinal(TemplateProfileRegion region)
    {
        return ParseBlockOrdinal(region.TemplateBlockId ?? region.StartBlockId) ?? int.MaxValue;
    }

    private static IReadOnlyList<Paragraph>? ResolveTargetParagraphSpan(
        TemplateProfileRegion region,
        IReadOnlyList<Paragraph> paragraphs,
        out string? error)
    {
        error = null;

        var singleBlockId = region.TemplateBlockId;
        if (!string.IsNullOrWhiteSpace(singleBlockId))
        {
            var ordinal = ParseBlockOrdinal(singleBlockId);
            if (ordinal is null || ordinal.Value < 1 || ordinal.Value > paragraphs.Count)
            {
                error = $"No template block found for region `{region.Role}`.";
                return null;
            }

            return [paragraphs[ordinal.Value - 1]];
        }

        var startOrdinal = ParseBlockOrdinal(region.StartBlockId);
        var endOrdinal = ParseBlockOrdinal(region.EndBlockId);
        if (startOrdinal is null || endOrdinal is null || startOrdinal.Value < 1 || endOrdinal.Value < 1 ||
            startOrdinal.Value > paragraphs.Count || endOrdinal.Value > paragraphs.Count || endOrdinal.Value < startOrdinal.Value)
        {
            error = $"No template block found for region `{region.Role}`.";
            return null;
        }

        return paragraphs.Skip(startOrdinal.Value - 1).Take(endOrdinal.Value - startOrdinal.Value + 1).ToArray();
    }

    private static int? ParseBlockOrdinal(string? blockId)
    {
        if (string.IsNullOrWhiteSpace(blockId))
        {
            return null;
        }

        var digits = new string(blockId.Where(char.IsDigit).ToArray());
        return int.TryParse(digits, out var parsed) ? parsed : null;
    }

    private static string ResolveReplacementText(TemplateProfileRegion region, TemplateInspectionDocument sourceInspection)
    {
        var candidates = sourceInspection.Candidates.ToArray();
        if (string.Equals(region.Role, "title", StringComparison.OrdinalIgnoreCase))
        {
            return candidates.FirstOrDefault(candidate =>
                    !string.IsNullOrWhiteSpace(candidate.Text) &&
                    !candidate.StructuralHints.LooksLikeReference)?.Text
                ?? candidates.FirstOrDefault(candidate => !candidate.StructuralHints.LooksLikeReference)?.Text
                ?? candidates.FirstOrDefault()?.Text
                ?? "";
        }

        if (string.Equals(region.Role, "abstract", StringComparison.OrdinalIgnoreCase))
        {
            var abstractHeading = Array.FindIndex(candidates, candidate => IsAbstractHeading(candidate.Text));
            if (abstractHeading >= 0 && abstractHeading + 1 < candidates.Length)
            {
                var abstractCandidate = candidates
                    .Skip(abstractHeading + 1)
                    .FirstOrDefault(candidate => !string.IsNullOrWhiteSpace(candidate.Text));
                if (abstractCandidate is not null)
                {
                    return abstractCandidate.Text;
                }
            }
        }

        if (string.Equals(region.Role, "references", StringComparison.OrdinalIgnoreCase))
        {
            return ResolveReferenceSection(sourceInspection).Reference?.Text ?? "";
        }

        return candidates.FirstOrDefault(candidate => !candidate.StructuralHints.ShortHighlightedParagraph && !candidate.StructuralHints.LooksLikeReference)?.Text ?? "";
    }

    private static void ReplaceParagraphTextPreservingFormat(Paragraph paragraph, string replacement)
    {
        var firstRun = paragraph.Elements<Run>().FirstOrDefault();
        ClearParagraphContent([paragraph]);
        var run = firstRun is null ? new Run() : (Run)firstRun.CloneNode(true);
        run.RemoveAllChildren<Text>();
        run.Append(new Text(replacement) { Space = DocumentFormat.OpenXml.SpaceProcessingModeValues.Preserve });
        paragraph.Append(run);
    }

    private static Paragraph? ResolveReferenceParagraph(TemplateInspectionDocument sourceInspection, Body sourceBody)
    {
        var referenceCandidate = ResolveReferenceSection(sourceInspection).Reference;
        if (referenceCandidate is null)
        {
            return null;
        }

        return GetParagraphByOrdinal(sourceBody, referenceCandidate.Ordinal);
    }

    private static (TemplateInspectionCandidate? LeadIn, TemplateInspectionCandidate? Heading, TemplateInspectionCandidate? Reference) ResolveReferenceSection(TemplateInspectionDocument sourceInspection)
    {
        var candidates = sourceInspection.Candidates.ToArray();
        var headingIndex = Array.FindIndex(candidates, candidate => IsReferencesHeading(candidate.Text));
        var referenceCandidates = headingIndex >= 0 ? candidates.Skip(headingIndex + 1) : candidates;

        var referenceCandidate = referenceCandidates.FirstOrDefault(IsNonEmptyReferenceCandidate)
            ?? candidates.FirstOrDefault(IsNonEmptyReferenceCandidate);

        return headingIndex >= 0
            ? (FindLeadInCandidate(candidates, headingIndex), candidates[headingIndex], referenceCandidate)
            : (null, null, referenceCandidate);
    }

    private static TemplateInspectionCandidate? FindLeadInCandidate(IReadOnlyList<TemplateInspectionCandidate> candidates, int headingIndex)
    {
        for (var index = headingIndex - 1; index >= 0; index--)
        {
            var candidate = candidates[index];
            if (!string.IsNullOrWhiteSpace(candidate.Text) && !candidate.StructuralHints.LooksLikeReference)
            {
                return candidate;
            }
        }

        return null;
    }

    private static bool IsNonEmptyReferenceCandidate(TemplateInspectionCandidate candidate)
    {
        return !string.IsNullOrWhiteSpace(candidate.Text) && candidate.StructuralHints.LooksLikeReference;
    }

    private static Paragraph? GetParagraphByOrdinal(Body body, int ordinal)
    {
        if (ordinal < 1)
        {
            return null;
        }

        return body.Descendants<Paragraph>().ElementAtOrDefault(ordinal - 1);
    }

    private static void ReplaceParagraphWithClone(Paragraph target, Paragraph source)
    {
        ClearParagraphContent([target]);
        target.RemoveAllChildren<ParagraphProperties>();

        if (source.ParagraphProperties is not null)
        {
            target.PrependChild((ParagraphProperties)source.ParagraphProperties.CloneNode(true));
        }

        foreach (var child in source.ChildElements)
        {
            if (child is ParagraphProperties)
            {
                continue;
            }

            target.Append((OpenXmlElement)child.CloneNode(true));
        }
    }

    private static void InsertParagraphAfter(Paragraph target, Paragraph source)
    {
        target.InsertAfterSelf((Paragraph)source.CloneNode(true));
    }

    private static void ClearParagraphContent(IEnumerable<Paragraph> paragraphs)
    {
        foreach (var paragraph in paragraphs)
        {
            var children = paragraph.ChildElements.Where(child => child is not ParagraphProperties).ToArray();
            foreach (var child in children)
            {
                child.Remove();
            }
        }
    }

    private static string GetParagraphText(Paragraph paragraph)
    {
        return string.Concat(paragraph.Descendants<Text>().Select(t => t.Text)).Trim();
    }

    private static bool IsPreservedRegionText(string text)
    {
        return !string.IsNullOrWhiteSpace(text) &&
               (text.Contains("PRESERVADA", StringComparison.OrdinalIgnoreCase) ||
                text.Contains("TEMPLATE", StringComparison.OrdinalIgnoreCase));
    }

    private static bool IsMissingRequiredRegionError(string error)
    {
        return error.Contains("regiao obrigatoria", StringComparison.OrdinalIgnoreCase) ||
               error.Contains("Required region missing", StringComparison.OrdinalIgnoreCase);
    }

    private static bool IsAbstractHeading(string text)
    {
        return text.Contains("RESUMO", StringComparison.OrdinalIgnoreCase) ||
               text.Contains("ABSTRACT", StringComparison.OrdinalIgnoreCase);
    }

    private static bool IsReferencesHeading(string text)
    {
        var trimmed = text.Trim();
        return string.Equals(trimmed, "REFERÊNCIAS", StringComparison.OrdinalIgnoreCase) ||
               string.Equals(trimmed, "REFERENCIAS", StringComparison.OrdinalIgnoreCase);
    }

    private static List<string> ValidateInputs(string templatePath, string sourcePath, string profilePath, string outputPath)
    {
        var errors = new List<string>();
        if (string.IsNullOrWhiteSpace(templatePath) || !File.Exists(Path.GetFullPath(templatePath)))
        {
            errors.Add($"Template not found: {templatePath}");
        }

        if (string.IsNullOrWhiteSpace(sourcePath) || !File.Exists(Path.GetFullPath(sourcePath)))
        {
            errors.Add($"Source not found: {sourcePath}");
        }

        if (string.IsNullOrWhiteSpace(profilePath) || !File.Exists(Path.GetFullPath(profilePath)))
        {
            errors.Add($"Profile not found: {profilePath}");
        }

        if (string.IsNullOrWhiteSpace(outputPath))
        {
            errors.Add("--out is required");
        }

        return errors;
    }

    private sealed record TemplateApplicationSummary(
        IReadOnlyList<string> AppliedRegions,
        IReadOnlyList<string> PreservedRegions,
        IReadOnlyList<string> PendingIssues);
}

internal sealed record TemplateApplicationResult(
    bool Success,
    int ExitCode,
    TemplateApplicationReport Report,
    IReadOnlyList<string> Errors);
