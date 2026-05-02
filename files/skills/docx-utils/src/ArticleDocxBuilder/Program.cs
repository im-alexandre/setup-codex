using System.Globalization;
using System.IO.Packaging;
using System.Text.Json;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using DocumentFormat.OpenXml.Wordprocessing;
using A = DocumentFormat.OpenXml.Drawing;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using PIC = DocumentFormat.OpenXml.Drawing.Pictures;

if (args.Length < 2 || args[0] is "-h" or "--help")
{
    Console.Error.WriteLine("Usage: ArticleDocxBuilder <article_spec.json> <output.docx> [author] [--lock <lockfile>] [--template <template.docx>] [--sbpo] [--blind]");
    return 2;
}

var specPath = Path.GetFullPath(args[0]);
var outputPath = Path.GetFullPath(args[1]);
var author = "Marina Saldanha";
var lockPath = "";
var templatePath = "";
var sbpoMode = false;
var blindMode = false;
for (var i = 2; i < args.Length; i++)
{
    if (args[i] == "--lock" && i + 1 < args.Length)
    {
        lockPath = Path.GetFullPath(args[++i]);
        continue;
    }
    if (args[i] == "--template" && i + 1 < args.Length)
    {
        templatePath = Path.GetFullPath(args[++i]);
        continue;
    }
    if (args[i] == "--sbpo")
    {
        sbpoMode = true;
        continue;
    }
    if (args[i] == "--blind")
    {
        blindMode = true;
        continue;
    }
    if (!args[i].StartsWith("--", StringComparison.Ordinal))
    {
        author = args[i];
    }
}

var json = File.ReadAllText(specPath);
var options = new JsonSerializerOptions { PropertyNameCaseInsensitive = true };
var spec = JsonSerializer.Deserialize<ArticleSpec>(json, options)
           ?? throw new InvalidOperationException("Could not parse article spec.");

Directory.CreateDirectory(Path.GetDirectoryName(outputPath)!);
using var lockStream = AcquireLock(lockPath);
if (File.Exists(outputPath))
{
    File.Delete(outputPath);
}

if (!string.IsNullOrWhiteSpace(templatePath))
{
    File.Copy(templatePath, outputPath);
}

using (var doc = string.IsNullOrWhiteSpace(templatePath)
           ? WordprocessingDocument.Create(outputPath, WordprocessingDocumentType.Document)
           : WordprocessingDocument.Open(outputPath, true))
{
    var main = doc.MainDocumentPart ?? doc.AddMainDocumentPart();
    var existingSectPr = main.Document?.Body?.Elements<SectionProperties>().LastOrDefault()?.CloneNode(true) as SectionProperties;
    main.Document = new Document(new Body());
    AddCoreProperties(doc, spec, author);
    AddSettings(main);
    AddStyles(main, sbpoMode);

    var body = main.Document.Body!;

    body.Append(CreateParagraph(spec.Title, "Title", JustificationValues.Center));
    if (!blindMode && !string.IsNullOrWhiteSpace(spec.Subtitle))
    {
        body.Append(CreateParagraph(spec.Subtitle, "Subtitle", JustificationValues.Center));
    }
    if (!blindMode && !string.IsNullOrWhiteSpace(spec.AuthorLine))
    {
        body.Append(CreateParagraph(spec.AuthorLine, "Normal", JustificationValues.Center));
    }

    AddFrontMatter(body, "Resumo", spec.Resumo, "Palavras-chave", spec.PalavrasChave);
    AddFrontMatter(body, "Abstract", spec.Abstract, "Keywords", spec.Keywords);

    foreach (var section in spec.Sections)
    {
        body.Append(CreateParagraph(section.Heading, section.Level <= 1 ? "Heading1" : "Heading2"));
        foreach (var item in section.Items)
        {
            AddItem(doc, main, body, item);
        }
        foreach (var paragraph in section.Paragraphs)
        {
            body.Append(CreateParagraph(paragraph, "Normal", JustificationValues.Both));
        }
    }

    if (spec.References.Count > 0)
    {
        body.Append(CreateParagraph("REFERÊNCIAS", "Heading1"));
        foreach (var reference in spec.References)
        {
            body.Append(CreateReferenceParagraph(reference));
        }
    }

    body.Append(NormalizeSectionProperties(existingSectPr, sbpoMode));

    main.Document.Save();
}

using (var doc = WordprocessingDocument.Open(outputPath, true))
{
    var validator = new OpenXmlValidator(FileFormatVersions.Office2019);
    var errors = validator.Validate(doc).ToList();
    Console.WriteLine(JsonSerializer.Serialize(new
    {
        output = outputPath,
        validation_errors = errors.Count,
        first_errors = errors.Take(10).Select(e => new { e.Description, e.Path?.XPath }).ToList()
    }, new JsonSerializerOptions { WriteIndented = true }));
}

return 0;

static FileStream? AcquireLock(string lockPath)
{
    if (string.IsNullOrWhiteSpace(lockPath)) return null;
    Directory.CreateDirectory(Path.GetDirectoryName(lockPath)!);
    return new FileStream(lockPath, FileMode.OpenOrCreate, FileAccess.ReadWrite, FileShare.None);
}

static void AddCoreProperties(WordprocessingDocument doc, ArticleSpec spec, string author)
{
    doc.PackageProperties.Creator = author;
    doc.PackageProperties.Title = spec.Title;
    doc.PackageProperties.Subject = spec.Subtitle;
    doc.PackageProperties.Created = DateTime.UtcNow;
    doc.PackageProperties.Modified = DateTime.UtcNow;
    doc.PackageProperties.Language = "pt-BR";
}

static void AddSettings(MainDocumentPart main)
{
    if (main.DocumentSettingsPart is not null)
    {
        main.DeletePart(main.DocumentSettingsPart);
    }
    var settings = main.AddNewPart<DocumentSettingsPart>();
    settings.Settings = new Settings(
        new TrackRevisions(),
        new UpdateFieldsOnOpen { Val = true },
        new Compatibility(new CompatibilitySetting
        {
            Name = CompatSettingNameValues.CompatibilityMode,
            Uri = "http://schemas.microsoft.com/office/word",
            Val = "15"
        })
    );
    settings.Settings.Save();
}

static void AddStyles(MainDocumentPart main, bool sbpoMode)
{
    if (main.StyleDefinitionsPart is not null)
    {
        main.DeletePart(main.StyleDefinitionsPart);
    }
    var stylesPart = main.AddNewPart<StyleDefinitionsPart>();
    var styles = new Styles();
    var bodySize = sbpoMode ? "22" : "24";
    var titleSize = sbpoMode ? "24" : "28";
    var smallSize = sbpoMode ? "20" : "22";
    var tableSize = sbpoMode ? "18" : "18";
    var normalLine = sbpoMode ? "240" : "360";
    var normalAfter = sbpoMode ? "0" : "120";
    styles.Append(new DocDefaults(
        new RunPropertiesDefault(new RunPropertiesBaseStyle(
            new RunFonts { Ascii = "Times New Roman", HighAnsi = "Times New Roman", ComplexScript = "Times New Roman" },
            new FontSize { Val = bodySize },
            new FontSizeComplexScript { Val = bodySize })),
        new ParagraphPropertiesDefault(new ParagraphPropertiesBaseStyle(
            new SpacingBetweenLines { After = normalAfter, Line = normalLine, LineRule = LineSpacingRuleValues.Auto }))));

    styles.Append(MakeParagraphStyle("Normal", "Normal", false, bodySize, null, JustificationValues.Both, after: normalAfter, line: normalLine));
    styles.Append(MakeParagraphStyle("Title", "Title", true, titleSize, null, JustificationValues.Center, before: sbpoMode ? "0" : "240", after: sbpoMode ? "120" : "240", line: normalLine));
    styles.Append(MakeParagraphStyle("Subtitle", "Subtitle", false, bodySize, null, JustificationValues.Center, after: sbpoMode ? "120" : "240", line: normalLine));
    styles.Append(MakeParagraphStyle("Heading1", "Heading 1", true, bodySize, "Normal", JustificationValues.Left, before: sbpoMode ? "180" : "360", after: sbpoMode ? "60" : "180", line: normalLine));
    styles.Append(MakeParagraphStyle("Heading2", "Heading 2", true, bodySize, "Normal", JustificationValues.Left, before: sbpoMode ? "120" : "240", after: sbpoMode ? "60" : "120", line: normalLine));
    styles.Append(MakeParagraphStyle("Tabela", "Tabela", false, smallSize, "Normal", JustificationValues.Left, before: "120", after: "40", line: normalLine));
    styles.Append(MakeParagraphStyle("Figura", "Figura", false, smallSize, "Normal", JustificationValues.Left, before: "120", after: "40", line: normalLine));
    styles.Append(MakeParagraphStyle("legenda0", "legenda", false, tableSize, "Normal", JustificationValues.Left, before: "0", after: "80", line: "240"));
    styles.Append(MakeParagraphStyle("dados", "dados", false, tableSize, "Normal", JustificationValues.Left, before: "0", after: "0", line: "240"));
    styles.Append(MakeParagraphStyle("equao", "equacao", false, smallSize, "Normal", JustificationValues.Left, before: "80", after: "80", line: normalLine));
    styles.Append(MakeTableStyle());
    stylesPart.Styles = styles;
    styles.Save(stylesPart);
}

static Style MakeParagraphStyle(
    string id,
    string name,
    bool bold,
    string size,
    string? basedOn,
    JustificationValues justification,
    string before = "0",
    string after = "120",
    string line = "360")
{
    var pPr = new StyleParagraphProperties(
        new SpacingBetweenLines { Before = before, After = after, Line = line, LineRule = LineSpacingRuleValues.Auto });
    if (id == "Normal")
    {
        pPr.Append(new Indentation { FirstLine = "708" });
    }
    if (id == "Title" || id == "Subtitle" || id.StartsWith("Heading", StringComparison.Ordinal))
    {
        pPr.Append(new Indentation { FirstLine = "0" });
    }
    pPr.Append(new Justification { Val = justification });

    var rPr = new StyleRunProperties(
        new RunFonts { Ascii = "Times New Roman", HighAnsi = "Times New Roman", ComplexScript = "Times New Roman" });
    if (bold) rPr.Append(new Bold(), new BoldComplexScript());
    rPr.Append(
        new FontSize { Val = size },
        new FontSizeComplexScript { Val = size });

    var style = new Style { Type = StyleValues.Paragraph, StyleId = id, CustomStyle = id is not "Normal" };
    style.Append(new StyleName { Val = name });
    if (basedOn is not null) style.Append(new BasedOn { Val = basedOn });
    if (id == "Normal") style.Append(new PrimaryStyle());
    style.Append(pPr, rPr);
    return style;
}

static Style MakeTableStyle()
{
    var style = new Style { Type = StyleValues.Table, StyleId = "tabelauerj", CustomStyle = true };
    style.Append(new StyleName { Val = "tabela_uerj" });
    style.Append(new TableProperties(
        new TableBorders(
            new TopBorder { Val = BorderValues.Single, Size = 8, Color = "000000" },
            new BottomBorder { Val = BorderValues.Single, Size = 8, Color = "000000" },
            new InsideHorizontalBorder { Val = BorderValues.Single, Size = 4, Color = "808080" })));
    return style;
}

static SectionProperties NormalizeSectionProperties(SectionProperties? templateSectPr, bool sbpoMode)
{
    if (templateSectPr is not null && !sbpoMode)
    {
        return templateSectPr;
    }
    if (!sbpoMode)
    {
        return CreateSectionProperties(false);
    }
    return CreateSectionProperties(true);
}

static SectionProperties CreateSectionProperties(bool sbpoMode)
{
    if (sbpoMode)
    {
        return new SectionProperties(
            new PageSize { Width = 11906, Height = 16838 },
            new PageMargin { Top = 1871, Right = 1644, Bottom = 1418, Left = 1644, Header = 708, Footer = 708, Gutter = 0 });
    }
    return new SectionProperties(
        new PageSize { Width = 11906, Height = 16838 },
        new PageMargin { Top = 1701, Right = 1134, Bottom = 1134, Left = 1701, Header = 708, Footer = 708, Gutter = 0 });
}

static void AddFrontMatter(Body body, string heading, string text, string keywordLabel, List<string> keywords)
{
    if (string.IsNullOrWhiteSpace(text)) return;
    body.Append(CreateParagraph(heading, "Heading1"));
    body.Append(CreateParagraph(text, "Normal", JustificationValues.Both));
    if (keywords.Count > 0)
    {
        body.Append(CreateParagraph($"{keywordLabel}: {string.Join("; ", keywords)}.", "Normal", JustificationValues.Left));
    }
}

static void AddItem(WordprocessingDocument doc, MainDocumentPart main, Body body, ArticleItem item)
{
    switch (item.Kind.ToLowerInvariant())
    {
        case "paragraph":
            body.Append(CreateParagraph(item.Text ?? "", item.StyleId ?? "Normal", JustificationValues.Both));
            break;
        case "heading2":
            body.Append(CreateParagraph(item.Text ?? "", "Heading2"));
            break;
        case "equation":
            body.Append(CreateParagraph(item.Text ?? "", "equao", JustificationValues.Left));
            break;
        case "table":
            body.Append(CreateParagraph(item.Caption ?? "", "Tabela", JustificationValues.Left));
            body.Append(CreateTable(item.Headers, item.Rows));
            if (!string.IsNullOrWhiteSpace(item.SourceText))
            {
                body.Append(CreateParagraph(item.SourceText!, "legenda0", JustificationValues.Left));
            }
            break;
        case "figure":
            body.Append(CreateFigureParagraph(doc, main, item));
            if (!string.IsNullOrWhiteSpace(item.SourceText))
            {
                body.Append(CreateParagraph(item.SourceText!, "legenda0", JustificationValues.Left));
            }
            break;
    }
}

static Paragraph CreateParagraph(string text, string styleId, JustificationValues? justification = null)
{
    var pPr = new ParagraphProperties(new ParagraphStyleId { Val = styleId });
    if (justification.HasValue) pPr.Append(new Justification { Val = justification.Value });
    var p = new Paragraph(pPr);
    foreach (var (segment, isBreak) in SplitBreaks(text))
    {
        if (isBreak)
        {
            p.Append(new Run(new Break()));
        }
        else if (segment.Length > 0)
        {
            p.Append(new Run(new Text(segment) { Space = SpaceProcessingModeValues.Preserve }));
        }
    }
    return p;
}

static IEnumerable<(string Text, bool IsBreak)> SplitBreaks(string text)
{
    var parts = (text ?? "").Split('\n');
    for (var i = 0; i < parts.Length; i++)
    {
        if (i > 0) yield return ("", true);
        yield return (parts[i], false);
    }
}

static Paragraph CreateReferenceParagraph(string text)
{
    var p = CreateParagraph(text, "Normal", JustificationValues.Left);
    var pPr = p.GetFirstChild<ParagraphProperties>()!;
    pPr.Indentation = new Indentation { Left = "708", Hanging = "708" };
    pPr.SpacingBetweenLines = new SpacingBetweenLines { Before = "0", After = "120", Line = "240", LineRule = LineSpacingRuleValues.Auto };
    return p;
}

static Table CreateTable(List<string> headers, List<List<string>> rows)
{
    var table = new Table();
    table.Append(new TableProperties(
        new TableStyle { Val = "tabelauerj" },
        new TableWidth { Width = "9000", Type = TableWidthUnitValues.Dxa },
        new TableJustification { Val = TableRowAlignmentValues.Center },
        new TableLayout { Type = TableLayoutValues.Fixed },
        new TableLook { FirstRow = true, NoHorizontalBand = false, NoVerticalBand = true }));

    var colCount = Math.Max(headers.Count, rows.Count == 0 ? headers.Count : rows.Max(r => r.Count));
    var widths = CalculateColumnWidths(headers, rows, colCount);
    var grid = new TableGrid();
    foreach (var width in widths)
    {
        grid.Append(new GridColumn { Width = width.ToString(CultureInfo.InvariantCulture) });
    }
    table.Append(grid);
    table.Append(CreateTableRow(headers, widths, true));
    foreach (var row in rows)
    {
        table.Append(CreateTableRow(row, widths, false));
    }
    return table;
}

static int[] CalculateColumnWidths(List<string> headers, List<List<string>> rows, int colCount)
{
    var weights = new double[colCount];
    for (var i = 0; i < colCount; i++)
    {
        weights[i] = i < headers.Count ? Math.Max(6, headers[i].Length) : 6;
    }
    foreach (var row in rows)
    {
        for (var i = 0; i < Math.Min(colCount, row.Count); i++)
        {
            weights[i] = Math.Max(weights[i], Math.Min(45, row[i]?.Length ?? 0));
        }
    }
    var total = weights.Sum();
    return weights.Select(w => Math.Max(900, (int)Math.Round(9000 * w / total))).ToArray();
}

static TableRow CreateTableRow(List<string> cells, int[] widths, bool header)
{
    var tr = new TableRow();
    if (header)
    {
        tr.Append(new TableRowProperties(new TableHeader()));
    }
    for (var i = 0; i < widths.Length; i++)
    {
        var text = i < cells.Count ? cells[i] : "";
        var tc = new TableCell();
        tc.Append(new TableCellProperties(
            new TableCellWidth { Width = widths[i].ToString(CultureInfo.InvariantCulture), Type = TableWidthUnitValues.Dxa },
            new TableCellVerticalAlignment { Val = TableVerticalAlignmentValues.Center }));
        var p = CreateParagraph(text, "dados", header ? JustificationValues.Center : JustificationValues.Left);
        if (header)
        {
            foreach (var run in p.Elements<Run>())
            {
                run.RunProperties ??= new RunProperties();
                run.RunProperties.Append(new Bold());
            }
        }
        tc.Append(p);
        tr.Append(tc);
    }
    return tr;
}

static Paragraph CreateFigureParagraph(WordprocessingDocument doc, MainDocumentPart main, ArticleItem item)
{
    var paragraph = CreateParagraph(item.Caption ?? "", "Figura", JustificationValues.Left);
    paragraph.Append(new Run(new Break()));

    var imagePath = Path.GetFullPath(item.ImagePath ?? throw new InvalidOperationException("Figure item missing imagePath."));
    var imagePart = main.AddImagePart(ImagePartType.Png);
    using (var stream = File.OpenRead(imagePath))
    {
        imagePart.FeedData(stream);
    }
    var relationshipId = main.GetIdOfPart(imagePart);
    var (pxW, pxH) = ReadPngSize(imagePath);
    var widthCm = item.WidthCm is > 0 ? item.WidthCm.Value : 14.0;
    var widthEmu = (long)Math.Round(widthCm * 360000.0);
    var heightEmu = (long)Math.Round(widthEmu * pxH / Math.Max(1.0, pxW));
    var drawing = CreateAnchoredDrawing(relationshipId, item.Caption ?? Path.GetFileName(imagePath), widthEmu, heightEmu);
    paragraph.Append(new Run(drawing));
    return paragraph;
}

static (double Width, double Height) ReadPngSize(string path)
{
    using var fs = File.OpenRead(path);
    var header = new byte[24];
    fs.ReadExactly(header);
    var width = (header[16] << 24) + (header[17] << 16) + (header[18] << 8) + header[19];
    var height = (header[20] << 24) + (header[21] << 16) + (header[22] << 8) + header[23];
    return (width, height);
}

static Drawing CreateAnchoredDrawing(string relationshipId, string name, long widthEmu, long heightEmu)
{
    var docPrId = (UInt32Value)(uint)Math.Abs(HashCode.Combine(relationshipId, name));
    var anchor = new DW.Anchor
    {
        DistanceFromTop = 0U,
        DistanceFromBottom = 0U,
        DistanceFromLeft = 0U,
        DistanceFromRight = 0U,
        SimplePos = false,
        RelativeHeight = 0U,
        BehindDoc = false,
        Locked = false,
        LayoutInCell = true,
        AllowOverlap = false
    };
    anchor.Append(
        new DW.SimplePosition { X = 0L, Y = 0L },
        new DW.HorizontalPosition(new DW.HorizontalAlignment("center")) { RelativeFrom = DW.HorizontalRelativePositionValues.Margin },
        new DW.VerticalPosition(new DW.PositionOffset("0")) { RelativeFrom = DW.VerticalRelativePositionValues.Paragraph },
        new DW.Extent { Cx = widthEmu, Cy = heightEmu },
        new DW.EffectExtent { LeftEdge = 0L, TopEdge = 0L, RightEdge = 0L, BottomEdge = 0L },
        new DW.WrapTopBottom(),
        new DW.DocProperties { Id = docPrId, Name = name, Description = name },
        new DW.NonVisualGraphicFrameDrawingProperties(new A.GraphicFrameLocks { NoChangeAspect = true }),
        new A.Graphic(
            new A.GraphicData(
                new PIC.Picture(
                    new PIC.NonVisualPictureProperties(
                        new PIC.NonVisualDrawingProperties { Id = 0U, Name = name, Description = name },
                        new PIC.NonVisualPictureDrawingProperties()),
                    new PIC.BlipFill(
                        new A.Blip { Embed = relationshipId },
                        new A.Stretch(new A.FillRectangle())),
                    new PIC.ShapeProperties(
                        new A.Transform2D(
                            new A.Offset { X = 0L, Y = 0L },
                            new A.Extents { Cx = widthEmu, Cy = heightEmu }),
                        new A.PresetGeometry(new A.AdjustValueList()) { Preset = A.ShapeTypeValues.Rectangle })))
            { Uri = "http://schemas.openxmlformats.org/drawingml/2006/picture" }));
    return new Drawing(anchor);
}

public sealed class ArticleSpec
{
    public string Title { get; set; } = "";
    public string Subtitle { get; set; } = "";
    public string AuthorLine { get; set; } = "";
    public string Resumo { get; set; } = "";
    public string Abstract { get; set; } = "";
    public List<string> PalavrasChave { get; set; } = [];
    public List<string> Keywords { get; set; } = [];
    public List<ArticleSection> Sections { get; set; } = [];
    public List<string> References { get; set; } = [];
}

public sealed class ArticleSection
{
    public string Heading { get; set; } = "";
    public int Level { get; set; } = 1;
    public List<string> Paragraphs { get; set; } = [];
    public List<ArticleItem> Items { get; set; } = [];
}

public sealed class ArticleItem
{
    public string Kind { get; set; } = "paragraph";
    public string? Text { get; set; }
    public string? StyleId { get; set; }
    public string? Caption { get; set; }
    public string? SourceText { get; set; }
    public string? ImagePath { get; set; }
    public double? WidthCm { get; set; }
    public List<string> Headers { get; set; } = [];
    public List<List<string>> Rows { get; set; } = [];
}
