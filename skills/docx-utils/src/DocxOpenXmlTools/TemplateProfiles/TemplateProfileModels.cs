namespace DocxOpenXmlTools.TemplateProfiles;

internal sealed record TemplateInspectionDocument
{
    public TemplateInspectionTemplate Template { get; init; } = new();

    public IReadOnlyList<TemplateInspectionCandidate> Candidates { get; init; } = [];
}

internal sealed record TemplateInspectionTemplate
{
    public string Name { get; init; } = "";

    public string SourceFile { get; init; } = "";

    public string Sha256 { get; init; } = "";

    public int ProfileVersion { get; init; }
}

internal sealed record TemplateInspectionCandidate
{
    public string Id { get; init; } = "";

    public string Text { get; init; } = "";

    public TemplateInspectionLocation Location { get; init; } = new();

    public int Ordinal { get; init; }

    public TemplateInspectionParagraphFormat ParagraphFormat { get; init; } = new();

    public IReadOnlyList<TemplateInspectionRun> Runs { get; init; } = [];

    public TemplateInspectionStructuralHints StructuralHints { get; init; } = new();
}

internal sealed record TemplateInspectionLocation
{
    public string Part { get; init; } = "body";

    public int Ordinal { get; init; }
}

internal sealed record TemplateInspectionParagraphFormat
{
    public string StyleId { get; init; } = "Normal";

    public string Alignment { get; init; } = "left";

    public int? SpacingBefore { get; init; }

    public int? SpacingAfter { get; init; }

    public int? HangingIndent { get; init; }
}

internal sealed record TemplateInspectionRun
{
    public string Text { get; init; } = "";

    public bool Bold { get; init; }

    public bool Italic { get; init; }

    public bool AllCaps { get; init; }

    public int? FontSize { get; init; }
}

internal sealed record TemplateInspectionStructuralHints
{
    public string? ManualNumbering { get; init; }

    public bool ShortHighlightedParagraph { get; init; }

    public bool LooksLikeReference { get; init; }
}

internal sealed record TemplateProfileDocument
{
    public TemplateProfileTemplate Template { get; init; } = new();

    public IReadOnlyList<TemplateProfileRegion> Regions { get; init; } = [];
}

internal sealed record TemplateProfileTemplate
{
    public string Name { get; init; } = "";

    public string SourceFile { get; init; } = "";

    public string Sha256 { get; init; } = "";

    public int ProfileVersion { get; init; }
}

internal sealed record TemplateProfileRegion
{
    public string Role { get; init; } = "";

    public string? TemplateBlockId { get; init; }

    public string? StartBlockId { get; init; }

    public string? EndBlockId { get; init; }

    public string? ReplaceWith { get; init; }

    public bool? PreserveFormatting { get; init; }

    public string? PreserveFormattingFrom { get; init; }

    public string? ReferenceFormattingProfile { get; init; }
}

internal sealed record TemplateApplicationReport
{
    public string Template { get; init; } = "";

    public string Source { get; init; } = "";

    public string Profile { get; init; } = "";

    public string Output { get; init; } = "";

    public IReadOnlyList<string> AppliedRegions { get; init; } = [];

    public IReadOnlyList<string> PreservedRegions { get; init; } = [];

    public IReadOnlyList<string> PendingIssues { get; init; } = [];
}

internal sealed record TemplateAuditReport
{
    public string Docx { get; init; } = "";

    public string Profile { get; init; } = "";

    public bool OpenXmlValid { get; init; }

    public IReadOnlyList<string> PendingIssues { get; init; } = [];
}
