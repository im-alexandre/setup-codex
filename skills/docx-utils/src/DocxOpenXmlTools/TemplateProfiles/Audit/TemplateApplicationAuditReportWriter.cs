using System.Text;

namespace DocxOpenXmlTools.TemplateProfiles.Audit;

internal static class TemplateApplicationAuditReportWriter
{
    public static string BuildMarkdown(TemplateAuditReport report)
    {
        var builder = new StringBuilder();
        builder.AppendLine("# Template Application Audit");
        builder.AppendLine();
        builder.AppendLine($"- DOCX: `{report.Docx}`");
        builder.AppendLine($"- Profile: `{report.Profile}`");
        builder.AppendLine($"- Open XML: {(report.OpenXmlValid ? "OK" : "INVALID")}");
        builder.AppendLine($"- Pendencias: {report.PendingIssues.Count}");
        builder.AppendLine();
        builder.AppendLine("## Pendencias");
        foreach (var issue in report.PendingIssues)
        {
            builder.AppendLine($"- {issue}");
        }

        return builder.ToString();
    }
}
