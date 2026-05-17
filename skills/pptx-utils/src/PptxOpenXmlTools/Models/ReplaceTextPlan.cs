namespace PptxOpenXmlTools.Models;

internal sealed class ReplaceTextPlan
{
    public List<ReplaceTextItem>? Replacements { get; set; }
}

internal sealed class ReplaceTextItem
{
    public int? SlideIndex { get; set; }
    public int? ShapeIndex { get; set; }
    public string? Find { get; set; }
    public string? Replace { get; set; }
    public string? Mode { get; set; }
}
