using System.Text;

namespace DocxOpenXmlTools.TemplateProfiles.Application;

internal static class TemplateApplicationReportWriter
{
    public static string BuildMarkdown(TemplateApplicationReport report)
    {
        var builder = new StringBuilder();
        builder.AppendLine("# Template Application Report");
        builder.AppendLine();
        builder.AppendLine($"- Template: `{report.Template}`");
        builder.AppendLine($"- Source: `{report.Source}`");
        builder.AppendLine($"- Profile: `{report.Profile}`");
        builder.AppendLine($"- Output: `{report.Output}`");
        builder.AppendLine($"- Regioes aplicadas: {string.Join(", ", report.AppliedRegions)}");
        builder.AppendLine($"- Regioes preservadas: {string.Join(", ", report.PreservedRegions)}");
        builder.AppendLine($"- Pendencias: {report.PendingIssues.Count}");
        builder.AppendLine();

        builder.AppendLine("## Regioes aplicadas");
        foreach (var region in report.AppliedRegions)
        {
            builder.AppendLine($"- {region}");
        }

        builder.AppendLine();
        builder.AppendLine("## Regioes preservadas");
        foreach (var region in report.PreservedRegions)
        {
            builder.AppendLine($"- {region}");
        }

        builder.AppendLine();
        builder.AppendLine("## Pendencias");
        foreach (var issue in report.PendingIssues)
        {
            builder.AppendLine($"- {issue}");
        }

        return builder.ToString();
    }
}
