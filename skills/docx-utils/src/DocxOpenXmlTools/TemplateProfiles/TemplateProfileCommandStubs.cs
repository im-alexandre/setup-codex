using System.Globalization;
using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using DocxOpenXmlTools.Cli;
using DocxOpenXmlTools.TemplateProfiles.Application;
using DocxOpenXmlTools.TemplateProfiles.Audit;
using DocxOpenXmlTools.TemplateProfiles.Inspection;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using DocumentFormat.OpenXml.Wordprocessing;

namespace DocxOpenXmlTools.TemplateProfiles;

internal static class TemplateProfileCommandStubs
{
    public static int InspectTemplate(string[] args)
    {
        var parseStatus = TryParsePositionalAndOptions(args, requirePositional: true, out var positional, out var options);
        if (parseStatus != 0)
        {
            PrintInspectUsage();
            return parseStatus;
        }

        var templatePath = ResolveExistingFile(positional!, isProfile: false);
        if (templatePath is null)
        {
            return 2;
        }

        if (!options.TryGetValue("out", out var outPathValue) || string.IsNullOrWhiteSpace(outPathValue))
        {
            Console.Error.WriteLine("--out is required");
            return 4;
        }

        var reportPath = options.TryGetValue("report", out var reportPathValue) && !string.IsNullOrWhiteSpace(reportPathValue)
            ? Path.GetFullPath(reportPathValue)
            : null;

        var inspection = TemplateDocumentInspector.Inspect(templatePath);
        var json = JsonSerializer.Serialize(inspection, WriteJsonOptions());
        var outputPath = Path.GetFullPath(outPathValue);
        Directory.CreateDirectory(Path.GetDirectoryName(outputPath) ?? ".");
        File.WriteAllText(outputPath, json, Encoding.UTF8);

        if (reportPath is not null)
        {
            Directory.CreateDirectory(Path.GetDirectoryName(reportPath) ?? ".");
            File.WriteAllText(reportPath, TemplateCandidateReportWriter.BuildMarkdown(inspection), Encoding.UTF8);
        }

        Console.WriteLine($"Template inspected: {templatePath}");
        return 0;
    }

    public static int ValidateTemplateProfile(string[] args)
    {
        var parseStatus = TryParsePositionalAndOptions(args, requirePositional: true, out var positional, out _);
        if (parseStatus != 0)
        {
            PrintValidateUsage();
            return parseStatus;
        }

        var profilePath = ResolveExistingFile(positional!, isProfile: true);
        if (profilePath is null)
        {
            return 5;
        }

        var validation = TemplateProfileValidator.Validate(profilePath);
        if (!validation.IsValid || validation.Profile is null)
        {
            foreach (var error in validation.Errors)
            {
                Console.Error.WriteLine(error);
            }

            return 6;
        }

        Console.WriteLine($"Profile valido: {profilePath}");
        Console.WriteLine($"Template: {validation.ResolvedTemplatePath}");
        Console.WriteLine($"Regioes: {validation.Profile.Regions.Count}");
        return 0;
    }

    public static int ApplyTemplate(string[] args)
    {
        var parseStatus = TryParsePositionalAndOptions(args, requirePositional: false, out _, out var options);
        if (parseStatus != 0)
        {
            PrintApplyUsage();
            return parseStatus;
        }

        if (!options.TryGetValue("template", out var templateValue) || string.IsNullOrWhiteSpace(templateValue))
        {
            Console.Error.WriteLine("--template is required");
            return 4;
        }

        if (!options.TryGetValue("source", out var sourceValue) || string.IsNullOrWhiteSpace(sourceValue))
        {
            Console.Error.WriteLine("--source is required");
            return 4;
        }

        if (!options.TryGetValue("profile", out var profileValue) || string.IsNullOrWhiteSpace(profileValue))
        {
            Console.Error.WriteLine("--profile is required");
            return 4;
        }

        if (!options.TryGetValue("out", out var outValue) || string.IsNullOrWhiteSpace(outValue))
        {
            Console.Error.WriteLine("--out is required");
            return 4;
        }

        var templatePath = Path.GetFullPath(templateValue);
        if (!File.Exists(templatePath))
        {
            Console.Error.WriteLine($"Template not found: {templatePath}");
            return 2;
        }

        var sourcePath = Path.GetFullPath(sourceValue);
        if (!File.Exists(sourcePath))
        {
            Console.Error.WriteLine($"Source not found: {sourcePath}");
            return 2;
        }

        var profilePath = Path.GetFullPath(profileValue);
        if (!File.Exists(profilePath))
        {
            Console.Error.WriteLine($"Profile not found: {profilePath}");
            return 5;
        }

        var outputPath = Path.GetFullPath(outValue);
        var reportPath = options.TryGetValue("report", out var reportValue) && !string.IsNullOrWhiteSpace(reportValue)
            ? Path.GetFullPath(reportValue)
            : null;

        var application = TemplateApplier.Apply(templatePath, sourcePath, profilePath, outputPath, reportPath);
        if (!application.Success)
        {
            foreach (var error in application.Errors)
            {
                Console.Error.WriteLine(error);
            }

            return application.ExitCode;
        }

        Console.WriteLine($"Template applied: {application.Report.Output}");
        Console.WriteLine($"Template profile: {application.Report.Profile}");
        return 0;
    }

    public static int AuditTemplateApplication(string[] args)
    {
        var parseStatus = TryParsePositionalAndOptions(args, requirePositional: true, out var positional, out var options);
        if (parseStatus != 0)
        {
            PrintAuditUsage();
            return parseStatus;
        }

        var docxPath = ResolveExistingFile(positional!, isProfile: false);
        if (docxPath is null)
        {
            return 2;
        }

        if (!options.TryGetValue("profile", out var profileValue) || string.IsNullOrWhiteSpace(profileValue))
        {
            Console.Error.WriteLine("--profile is required");
            return 4;
        }

        var profilePath = Path.GetFullPath(profileValue);
        if (!File.Exists(profilePath))
        {
            Console.Error.WriteLine($"Profile not found: {profilePath}");
            return 5;
        }

        var reportPath = options.TryGetValue("report", out var reportValue) && !string.IsNullOrWhiteSpace(reportValue)
            ? Path.GetFullPath(reportValue)
            : null;

        var audit = TemplateApplicationAuditor.Audit(docxPath, profilePath, reportPath);
        if (!audit.Success)
        {
            foreach (var issue in audit.Errors)
            {
                Console.Error.WriteLine(issue);
            }

            return audit.ExitCode;
        }

        Console.WriteLine($"Audit complete: {docxPath}");
        Console.WriteLine(audit.Report.OpenXmlValid ? "Open XML valid" : "Open XML invalid");
        return audit.ExitCode;
    }

    private static TemplateInspectionDocument BuildInspectionDocument(string templatePath)
    {
        using var document = WordprocessingDocument.Open(templatePath, false);
        var mainPart = document.MainDocumentPart;
        var paragraphs = mainPart?.Document?.Body?.Elements<Paragraph>().ToList() ?? [];
        var sha256 = ComputeSha256(templatePath);

        return new TemplateInspectionDocument
        {
            Template = new TemplateInspectionTemplate
            {
                Name = Path.GetFileNameWithoutExtension(templatePath),
                SourceFile = Path.GetFileName(templatePath),
                Sha256 = sha256,
                ProfileVersion = 1
            },
            Candidates = paragraphs.Select((paragraph, index) => BuildCandidate(paragraph, index + 1)).ToArray()
        };
    }

    private static TemplateInspectionCandidate BuildCandidate(Paragraph paragraph, int ordinal)
    {
        var text = string.Concat(paragraph.Descendants<Text>().Select(t => t.Text));
        var paragraphProperties = paragraph.ParagraphProperties;
        var styleId = paragraphProperties?.ParagraphStyleId?.Val?.Value ?? "Normal";
        var alignment = paragraphProperties?.Justification?.Val?.Value.ToString().ToLowerInvariant() ?? "left";
        var spacingBefore = TryParseInt(paragraphProperties?.SpacingBetweenLines?.Before?.Value);
        var spacingAfter = TryParseInt(paragraphProperties?.SpacingBetweenLines?.After?.Value);
        var hangingIndent = TryParseInt(paragraphProperties?.Indentation?.Hanging?.Value);
        var runs = paragraph.Elements<Run>().Select(BuildRun).ToArray();
        var manualNumbering = TryGetManualNumbering(text);

        return new TemplateInspectionCandidate
        {
            Id = $"p-{ordinal:0000}",
            Text = text,
            Ordinal = ordinal,
            Location = new TemplateInspectionLocation
            {
                Part = "body",
                Ordinal = ordinal
            },
            ParagraphFormat = new TemplateInspectionParagraphFormat
            {
                StyleId = styleId,
                Alignment = alignment,
                SpacingBefore = spacingBefore,
                SpacingAfter = spacingAfter,
                HangingIndent = hangingIndent
            },
            Runs = runs,
            StructuralHints = new TemplateInspectionStructuralHints
            {
                ManualNumbering = manualNumbering,
                ShortHighlightedParagraph = text.Length <= 40 && runs.Any(run => run.Bold || run.Italic || run.AllCaps),
                LooksLikeReference = text.Contains("doi", StringComparison.OrdinalIgnoreCase) ||
                    text.Contains("http", StringComparison.OrdinalIgnoreCase) ||
                    hangingIndent.HasValue
            }
        };
    }

    private static TemplateInspectionRun BuildRun(Run run)
    {
        var text = string.Concat(run.Descendants<Text>().Select(t => t.Text));
        var runProperties = run.RunProperties;
        return new TemplateInspectionRun
        {
            Text = text,
            Bold = runProperties?.Bold is not null,
            Italic = runProperties?.Italic is not null,
            AllCaps = runProperties?.Caps is not null || runProperties?.SmallCaps is not null,
            FontSize = TryParseInt(runProperties?.FontSize?.Val?.Value)
        };
    }

    private static string BuildInspectionReport(string templatePath, TemplateInspectionDocument inspection)
    {
        var report = new StringBuilder();
        report.AppendLine("# Template Inspection Report");
        report.AppendLine();
        report.AppendLine($"- Template: `{templatePath}`");
        report.AppendLine($"- Style predominante: {inspection.Candidates.FirstOrDefault()?.ParagraphFormat.StyleId ?? "Normal"}");
        report.AppendLine($"- Candidatos extraidos: {inspection.Candidates.Count}");
        report.AppendLine($"- Referencia detectada: {inspection.Candidates.Count(candidate => candidate.StructuralHints.LooksLikeReference)}");
        report.AppendLine();
        foreach (var candidate in inspection.Candidates)
        {
            report.AppendLine($"- {candidate.Id}: {candidate.Text}");
        }

        return report.ToString();
    }

    private static bool ValidateProfileDocument(
        string profilePath,
        out TemplateProfileDocument profile,
        out string resolvedTemplatePath,
        out List<string> errors)
    {
        errors = [];
        try
        {
            profile = JsonSerializer.Deserialize<TemplateProfileDocument>(File.ReadAllText(profilePath, Encoding.UTF8), ReadJsonOptions()) ?? new TemplateProfileDocument();
        }
        catch (JsonException ex)
        {
            errors.Add($"Profile invalido: nao foi possivel ler JSON ({ex.Message}).");
            profile = new TemplateProfileDocument();
            resolvedTemplatePath = profilePath;
            return false;
        }

        profile = profile with
        {
            Template = profile.Template ?? new TemplateProfileTemplate(),
            Regions = profile.Regions ?? []
        };
        resolvedTemplatePath = ResolveProfileTemplatePath(profilePath, profile.Template.SourceFile);

        if (string.IsNullOrWhiteSpace(profile.Template.Name))
        {
            errors.Add("Profile invalido: template.name ausente.");
        }

        if (string.IsNullOrWhiteSpace(profile.Template.SourceFile))
        {
            errors.Add("Profile invalido: template.sourceFile ausente.");
        }

        if (string.IsNullOrWhiteSpace(profile.Template.Sha256))
        {
            errors.Add("Profile invalido: template.sha256 ausente.");
        }

        if (profile.Template.ProfileVersion <= 0)
        {
            errors.Add("Profile invalido: template.profileVersion deve ser maior que zero.");
        }

        var requiredRegions = new[] { "title", "abstract", "references" };
        foreach (var requiredRegion in requiredRegions)
        {
            if (!profile.Regions.Any(region => string.Equals(region.Role, requiredRegion, StringComparison.OrdinalIgnoreCase)))
            {
                errors.Add($"Required region missing: {requiredRegion}");
            }
        }

        if (!File.Exists(resolvedTemplatePath))
        {
            errors.Add($"Template not found: {resolvedTemplatePath}");
            return false;
        }

        var actualHash = ComputeSha256(resolvedTemplatePath);
        if (!string.Equals(actualHash, profile.Template.Sha256, StringComparison.OrdinalIgnoreCase))
        {
            errors.Add($"Perfil invalido: hash/sha256 divergente. Esperado {profile.Template.Sha256}, encontrado {actualHash}.");
        }

        return errors.Count == 0;
    }

    private static IEnumerable<string> FindPendingIssues(string docxPath)
    {
        using var document = WordprocessingDocument.Open(docxPath, false);
        var body = document.MainDocumentPart?.Document?.Body;
        if (body is null)
        {
            yield return "Regiao obrigatoria ausente: body vazio.";
            yield break;
        }

        foreach (var paragraph in body.Elements<Paragraph>())
        {
            var text = string.Concat(paragraph.Descendants<Text>().Select(t => t.Text));
            if (text.Contains("REGIAO PRESERVADA DO TEMPLATE", StringComparison.OrdinalIgnoreCase) ||
                text.Contains("Conteudo exemplo do template", StringComparison.OrdinalIgnoreCase) ||
                text.Contains("Conteudo exemplo do resumo", StringComparison.OrdinalIgnoreCase))
            {
                yield return $"Regiao obrigatoria ausente: {text}";
            }
        }
    }

    private static string BuildApplicationReport(TemplateApplicationReport report)
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
        builder.AppendLine("## Pendencias");
        foreach (var issue in report.PendingIssues)
        {
            builder.AppendLine($"- {issue}");
        }

        return builder.ToString();
    }

    private static string BuildAuditReport(TemplateAuditReport report)
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

    private static int TryParsePositionalAndOptions(string[] args, bool requirePositional, out string? positional, out Dictionary<string, string> options)
    {
        positional = null;
        options = [];

        if (args.Length == 0 || CliOptions.IsHelpArgument(args[0]))
        {
            return 1;
        }

        if (requirePositional)
        {
            if (args.Length == 0 || args[0].StartsWith("-", StringComparison.Ordinal))
            {
                return 1;
            }

            positional = args[0];
            try
            {
                options = CliOptions.Parse(args.Skip(1).ToArray());
            }
            catch (ArgumentException ex)
            {
                Console.Error.WriteLine(ex.Message);
                return 4;
            }
        }
        else
        {
            try
            {
                options = CliOptions.Parse(args);
            }
            catch (ArgumentException ex)
            {
                Console.Error.WriteLine(ex.Message);
                return 4;
            }
        }

        return 0;
    }

    private static string? ResolveExistingFile(string path, bool isProfile)
    {
        var fullPath = Path.GetFullPath(path);
        if (!File.Exists(fullPath))
        {
            Console.Error.WriteLine($"{(isProfile ? "Profile" : "DOCX")} not found: {fullPath}");
            return null;
        }

        return fullPath;
    }

    private static string ResolveProfileTemplatePath(string profilePath, string templateSourceFile)
    {
        var baseDirectory = Path.GetDirectoryName(Path.GetFullPath(profilePath)) ?? ".";
        return Path.GetFullPath(Path.Combine(baseDirectory, templateSourceFile));
    }

    private static string ComputeSha256(string filePath)
    {
        using var stream = File.OpenRead(filePath);
        var hash = SHA256.HashData(stream);
        return Convert.ToHexString(hash).ToLowerInvariant();
    }

    private static int? TryParseInt(string? value)
    {
        if (int.TryParse(value, NumberStyles.Integer, CultureInfo.InvariantCulture, out var parsed))
        {
            return parsed;
        }

        return null;
    }

    private static string? TryGetManualNumbering(string text)
    {
        var index = 0;
        while (index < text.Length && char.IsWhiteSpace(text[index]))
        {
            index++;
        }

        var start = index;
        while (index < text.Length && char.IsDigit(text[index]))
        {
            index++;
        }

        if (index > start && index < text.Length && char.IsWhiteSpace(text[index]))
        {
            return text[start..index].Trim();
        }

        return null;
    }

    private static JsonSerializerOptions ReadJsonOptions() => new()
    {
        PropertyNameCaseInsensitive = true,
        ReadCommentHandling = JsonCommentHandling.Skip,
        AllowTrailingCommas = true
    };

    private static JsonSerializerOptions WriteJsonOptions() => new(CliOptions.JsonOptionsIndented())
    {
        PropertyNamingPolicy = JsonNamingPolicy.CamelCase
    };

    private static void PrintInspectUsage()
    {
        Console.WriteLine("Uso: docx-utils inspect-template <docx> --out <json> [--report <md>]");
    }

    private static void PrintValidateUsage()
    {
        Console.WriteLine("Uso: docx-utils validate-template-profile <profile.json>");
    }

    private static void PrintApplyUsage()
    {
        Console.WriteLine("Uso: docx-utils apply-template --template <docx> --source <docx> --profile <json> --out <docx> [--report <md>]");
    }

    private static void PrintAuditUsage()
    {
        Console.WriteLine("Uso: docx-utils audit-template-application <docx> --profile <json> [--report <md>]");
    }
}
