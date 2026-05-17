using System.Text;
using DocxOpenXmlTools.TemplateProfiles;

namespace DocxOpenXmlTools.TemplateProfiles.Inspection;

internal static class TemplateCandidateReportWriter
{
    public static string BuildMarkdown(TemplateInspectionDocument inspection)
    {
        ArgumentNullException.ThrowIfNull(inspection);

        var builder = new StringBuilder();
        var candidates = inspection.Candidates.ToArray();

        builder.AppendLine("# Template Inspection Report");
        builder.AppendLine();
        builder.AppendLine($"- Template: `{inspection.Template.Name}`");
        builder.AppendLine($"- Source file: `{inspection.Template.SourceFile}`");
        builder.AppendLine($"- Sha256: `{inspection.Template.Sha256}`");
        builder.AppendLine($"- Profile version: `{inspection.Template.ProfileVersion}`");
        builder.AppendLine($"- Candidates: {candidates.Length}");
        builder.AppendLine($"- Body/part counts: {string.Join(", ", candidates.GroupBy(candidate => candidate.Location.Part).Select(group => $"{group.Key}={group.Count()}"))}");
        builder.AppendLine();

        builder.AppendLine("## Summary");
        builder.AppendLine();
        builder.AppendLine("| Id | Part | Text | Style | Align | Hints |");
        builder.AppendLine("| --- | --- | --- | --- | --- | --- |");
        foreach (var candidate in candidates)
        {
            builder.Append('|')
                .Append(Escape(candidate.Id)).Append(" | ")
                .Append(Escape(candidate.Location.Part)).Append(" | ")
                .Append(Escape(Compact(candidate.Text))).Append(" | ")
                .Append(Escape(candidate.ParagraphFormat.StyleId)).Append(" | ")
                .Append(Escape(candidate.ParagraphFormat.Alignment)).Append(" | ")
                .Append(Escape(BuildHints(candidate))).AppendLine(" |");
        }

        builder.AppendLine();
        builder.AppendLine("## Candidate Details");
        builder.AppendLine();

        for (var i = 0; i < candidates.Length; i++)
        {
            var candidate = candidates[i];
            var previous = i > 0 ? candidates[i - 1] : null;
            var next = i + 1 < candidates.Length ? candidates[i + 1] : null;

            builder.AppendLine($"### {candidate.Id}");
            builder.AppendLine();
            builder.AppendLine($"- Part: `{candidate.Location.Part}`");
            builder.AppendLine($"- Ordinal: `{candidate.Ordinal}`");
            builder.AppendLine($"- Text: {FormatInline(candidate.Text)}");
            builder.AppendLine($"- Style: `{candidate.ParagraphFormat.StyleId}`");
            builder.AppendLine($"- Alignment: `{candidate.ParagraphFormat.Alignment}`");
            builder.AppendLine($"- Spacing before/after: `{FormatNullable(candidate.ParagraphFormat.SpacingBefore)}` / `{FormatNullable(candidate.ParagraphFormat.SpacingAfter)}`");
            builder.AppendLine($"- Hanging indent: `{FormatNullable(candidate.ParagraphFormat.HangingIndent)}`");
            builder.AppendLine($"- Manual numbering: `{candidate.StructuralHints.ManualNumbering ?? ""}`");
            builder.AppendLine($"- Short highlighted: `{candidate.StructuralHints.ShortHighlightedParagraph}`");
            builder.AppendLine($"- Looks like reference: `{candidate.StructuralHints.LooksLikeReference}`");

            if (previous is not null || next is not null)
            {
                builder.AppendLine($"- Neighbors: `{previous?.Id ?? ""}` <- `{candidate.Id}` -> `{next?.Id ?? ""}`");
            }

            builder.AppendLine("- Runs:");
            foreach (var run in candidate.Runs)
            {
                builder.AppendLine($"  - `{Compact(run.Text)}` bold=`{run.Bold}` italic=`{run.Italic}` allCaps=`{run.AllCaps}` fontSize=`{FormatNullable(run.FontSize)}`");
            }

            builder.AppendLine();
        }

        var referenceCandidates = candidates.Where(candidate => candidate.StructuralHints.LooksLikeReference).ToArray();
        if (referenceCandidates.Length > 0)
        {
            builder.AppendLine("## Candidatos de referencia");
            builder.AppendLine();
            foreach (var candidate in referenceCandidates)
            {
                builder.AppendLine($"- `{candidate.Id}` {FormatInline(candidate.Text)}");
            }
        }

        return builder.ToString();
    }

    private static string BuildHints(TemplateInspectionCandidate candidate)
    {
        var hints = new List<string>();
        if (!string.IsNullOrWhiteSpace(candidate.StructuralHints.ManualNumbering))
        {
            hints.Add($"manual:{candidate.StructuralHints.ManualNumbering}");
        }

        if (candidate.StructuralHints.ShortHighlightedParagraph)
        {
            hints.Add("short-highlight");
        }

        if (candidate.StructuralHints.LooksLikeReference)
        {
            hints.Add("reference");
        }

        return string.Join(", ", hints);
    }

    private static string Escape(string value) =>
        value.Replace("|", "\\|", StringComparison.Ordinal);

    private static string Compact(string value)
    {
        var trimmed = value.Trim();
        return trimmed.Length <= 80 ? trimmed : trimmed[..77] + "...";
    }

    private static string FormatInline(string value) => $"`{Escape(Compact(value))}`";

    private static string FormatNullable(int? value) => value?.ToString() ?? "";
}
