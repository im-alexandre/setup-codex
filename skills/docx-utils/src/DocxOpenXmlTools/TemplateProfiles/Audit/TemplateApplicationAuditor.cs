using System.Text;
using DocxOpenXmlTools.TemplateProfiles.Inspection;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using DocumentFormat.OpenXml.Wordprocessing;

namespace DocxOpenXmlTools.TemplateProfiles.Audit;

internal static class TemplateApplicationAuditor
{
    public static TemplateApplicationAuditResult Audit(string docxPath, string profilePath, string? reportPath)
    {
        var errors = new List<string>();
        if (string.IsNullOrWhiteSpace(docxPath) || !File.Exists(Path.GetFullPath(docxPath)))
        {
            errors.Add($"DOCX not found: {docxPath}");
        }

        if (string.IsNullOrWhiteSpace(profilePath) || !File.Exists(Path.GetFullPath(profilePath)))
        {
            errors.Add($"Profile not found: {profilePath}");
        }

        if (errors.Count > 0)
        {
            var missingReport = new TemplateAuditReport
            {
                Docx = docxPath,
                Profile = profilePath,
                OpenXmlValid = false,
                PendingIssues = errors
            };
            return new TemplateApplicationAuditResult(false, 4, missingReport, errors);
        }

        var validation = TemplateProfileValidator.Validate(profilePath);
        var pending = new List<string>();
        if (!validation.IsValid)
        {
            pending.AddRange(validation.Errors);
        }

        var templateInspection = validation.Profile is not null && !string.IsNullOrWhiteSpace(validation.ResolvedTemplatePath)
            ? TemplateDocumentInspector.Inspect(validation.ResolvedTemplatePath)
            : null;

        var ignoreBenignSpacingError = templateInspection is not null && !HasCanonicalReferenceTail(templateInspection);
        var openXmlValid = ValidateOpenXml(docxPath, pending, ignoreBenignSpacingError);
        if (validation.Profile is not null && !string.IsNullOrWhiteSpace(validation.ResolvedTemplatePath))
        {
            pending.AddRange(FindPendingIssues(docxPath, validation.ResolvedTemplatePath, validation.Profile));
        }

        var report = new TemplateAuditReport
        {
            Docx = Path.GetFullPath(docxPath),
            Profile = Path.GetFullPath(profilePath),
            OpenXmlValid = openXmlValid,
            PendingIssues = pending
        };

        if (!string.IsNullOrWhiteSpace(reportPath))
        {
            var normalizedReportPath = Path.GetFullPath(reportPath);
            Directory.CreateDirectory(Path.GetDirectoryName(normalizedReportPath) ?? ".");
            File.WriteAllText(normalizedReportPath, TemplateApplicationAuditReportWriter.BuildMarkdown(report), Encoding.UTF8);
        }

        var success = pending.Count == 0;
        return new TemplateApplicationAuditResult(success, success ? 0 : 6, report, pending);
    }

    private static bool ValidateOpenXml(string docxPath, List<string> pending, bool ignoreBenignSpacingError)
    {
        using var document = WordprocessingDocument.Open(Path.GetFullPath(docxPath), false);
        var errors = new OpenXmlValidator(DocumentFormat.OpenXml.FileFormatVersions.Office2019)
            .Validate(document)
            .Take(10)
            .ToArray();

        foreach (var error in errors)
        {
            if (ignoreBenignSpacingError && IsBenignStylesSpacingError(error.Description))
            {
                continue;
            }

            pending.Add($"Open XML validation: {error.Description}");
        }

        return errors.All(error => ignoreBenignSpacingError && IsBenignStylesSpacingError(error.Description));
    }

    private static bool IsBenignStylesSpacingError(string description)
    {
        return description.Contains("unexpected child element", StringComparison.OrdinalIgnoreCase) &&
               description.Contains(":spacing", StringComparison.OrdinalIgnoreCase);
    }

    private static bool HasCanonicalReferenceTail(TemplateInspectionDocument templateInspection)
    {
        return templateInspection.Candidates.LastOrDefault() is { } candidate &&
               IsCanonicalReferenceTail(candidate.Text);
    }

    private static bool IsCanonicalReferenceTail(string text)
    {
        return text.Contains("REGIAO PRESERVADA DO TEMPLATE", StringComparison.OrdinalIgnoreCase) ||
               text.Contains("Conteudo exemplo do resumo", StringComparison.OrdinalIgnoreCase);
    }

    private static IEnumerable<string> FindPendingIssues(string docxPath, string templatePath, TemplateProfileDocument profile)
    {
        var appliedInspection = TemplateDocumentInspector.Inspect(docxPath);
        var templateInspection = TemplateDocumentInspector.Inspect(templatePath);

        foreach (var region in profile.Regions)
        {
            var templateSpan = ResolveRegionSpan(region, templateInspection.Candidates, out var templateError);
            if (templateSpan is null)
            {
                yield return templateError ?? $"Regiao `{region.Role}` nao resolvida no template.";
                continue;
            }

            var appliedSpan = ResolveRegionSpan(region, appliedInspection.Candidates, out var appliedError);
            if (appliedSpan is null)
            {
                yield return appliedError ?? $"Regiao `{region.Role}` nao resolvida no DOCX aplicado.";
                continue;
            }

            if (IsReferenceRegion(region))
            {
                if (appliedSpan.Count != templateSpan.Count)
                {
                    yield return $"Regiao `{region.Role}` ficou incompleta no intervalo `{DescribeRegion(region)}`.";
                    continue;
                }

                if (appliedSpan[0].StructuralHints.LooksLikeReference &&
                    appliedSpan[0].Runs.Count < templateSpan[0].Runs.Count)
                {
                    yield return $"Regiao `{region.Role}` nao preservou runs de referencia em `{templateSpan[0].Id}`.";
                }

                for (var index = 1; index < appliedSpan.Count; index++)
                {
                    if (string.IsNullOrWhiteSpace(appliedSpan[index].Text) ||
                        IsReferencesHeading(appliedSpan[index].Text) ||
                        appliedSpan[index].StructuralHints.LooksLikeReference)
                    {
                        continue;
                    }

                    yield return $"Regiao `{region.Role}` ainda preserva texto residual em `{appliedSpan[index].Id}`.";
                }

                continue;
            }

            for (var index = 0; index < templateSpan.Count; index++)
            {
                if (index >= appliedSpan.Count)
                {
                    yield return $"Regiao `{region.Role}` ficou incompleta no intervalo `{DescribeRegion(region)}`.";
                    break;
                }

                if (string.IsNullOrWhiteSpace(templateSpan[index].Text) && string.IsNullOrWhiteSpace(appliedSpan[index].Text))
                {
                    continue;
                }

                if (string.Equals(appliedSpan[index].Text, templateSpan[index].Text, StringComparison.Ordinal))
                {
                    yield return $"Regiao `{region.Role}` ainda preserva texto do template em `{templateSpan[index].Id}`.";
                }
            }
        }

        foreach (var paragraph in appliedInspection.Candidates)
        {
            if (paragraph.Text.Contains("REGIAO PRESERVADA DO TEMPLATE", StringComparison.OrdinalIgnoreCase) ||
                paragraph.Text.Contains("Conteudo exemplo do template", StringComparison.OrdinalIgnoreCase) ||
                paragraph.Text.Contains("Conteudo exemplo do resumo", StringComparison.OrdinalIgnoreCase))
            {
                yield return $"Regiao obrigatoria pendente: {paragraph.Text}";
            }
        }
    }

    private static IReadOnlyList<TemplateInspectionCandidate>? ResolveRegionSpan(
        TemplateProfileRegion region,
        IReadOnlyList<TemplateInspectionCandidate> candidates,
        out string? error)
    {
        error = null;

        if (!string.IsNullOrWhiteSpace(region.TemplateBlockId))
        {
            var ordinal = ParseBlockOrdinal(region.TemplateBlockId);
            if (ordinal is null || ordinal.Value < 1 || ordinal.Value > candidates.Count)
            {
                error = $"Regiao `{region.Role}` nao resolvida para o bloco `{DescribeRegion(region)}`.";
                return null;
            }

            return [candidates[ordinal.Value - 1]];
        }

        var startOrdinal = ParseBlockOrdinal(region.StartBlockId);
        var endOrdinal = ParseBlockOrdinal(region.EndBlockId);
        if (startOrdinal is null || endOrdinal is null || startOrdinal.Value < 1 || endOrdinal.Value < 1 ||
            startOrdinal.Value > candidates.Count || endOrdinal.Value > candidates.Count || endOrdinal.Value < startOrdinal.Value)
        {
            error = $"Regiao `{region.Role}` nao resolvida para o intervalo `{DescribeRegion(region)}`.";
            return null;
        }

        return candidates.Skip(startOrdinal.Value - 1).Take(endOrdinal.Value - startOrdinal.Value + 1).ToArray();
    }

    private static bool IsReferenceRegion(TemplateProfileRegion region)
    {
        return string.Equals(region.Role, "references", StringComparison.OrdinalIgnoreCase) ||
               !string.IsNullOrWhiteSpace(region.ReferenceFormattingProfile);
    }

    private static bool IsReferencesHeading(string text)
    {
        var trimmed = text.Trim();
        return string.Equals(trimmed, "REFERÊNCIAS", StringComparison.OrdinalIgnoreCase) ||
               string.Equals(trimmed, "REFERENCIAS", StringComparison.OrdinalIgnoreCase);
    }

    private static string DescribeRegion(TemplateProfileRegion region)
    {
        if (!string.IsNullOrWhiteSpace(region.TemplateBlockId))
        {
            return region.TemplateBlockId;
        }

        return $"{region.StartBlockId ?? "?"}..{region.EndBlockId ?? "?"}";
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
}

internal sealed record TemplateApplicationAuditResult(
    bool Success,
    int ExitCode,
    TemplateAuditReport Report,
    IReadOnlyList<string> Errors);
