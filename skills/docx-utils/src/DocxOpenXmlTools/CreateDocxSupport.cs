using System.Text;
using System.Text.Json;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using DocumentFormat.OpenXml.Wordprocessing;

internal static class CreateDocxSupport
{
    public static int CreateDocx(string outputPath, IReadOnlyDictionary<string, string> options)
    {
        var planPath = options.TryGetValue("plan", out var planPathValue) && !string.IsNullOrWhiteSpace(planPathValue)
            ? Path.GetFullPath(planPathValue)
            : null;

        CreateDocxPlan? plan = null;
        if (planPath is not null)
        {
            if (!File.Exists(planPath))
            {
                Console.Error.WriteLine($"Plan not found: {planPath}");
                return 5;
            }

            var planJson = File.ReadAllText(planPath, Encoding.UTF8);
            var validation = PlanContractSupport.ValidateCreateDocxPlan(planJson);
            if (!validation.IsValid)
            {
                foreach (var error in validation.Errors)
                {
                    Console.Error.WriteLine(error);
                }
                return 6;
            }

            plan = JsonSerializer.Deserialize<CreateDocxPlan>(planJson, JsonOptions())
                ?? throw new InvalidOperationException("Could not parse create-docx plan.");
        }

        Directory.CreateDirectory(Path.GetDirectoryName(outputPath) ?? ".");
        if (File.Exists(outputPath))
        {
            File.Delete(outputPath);
        }

        using var doc = WordprocessingDocument.Create(outputPath, WordprocessingDocumentType.Document);
        var main = doc.AddMainDocumentPart();
        main.Document = new Document(new Body());
        main.Document.Body!.Append(new Paragraph());

        if (plan is not null)
        {
            RenderPlan(main.Document.Body!, plan);
        }

        main.Document.Save();

        var validationErrors = new OpenXmlValidator(FileFormatVersions.Office2019)
            .Validate(doc)
            .Take(10)
            .Select(e => new { e.Description, e.Path?.XPath })
            .ToList();

        Console.WriteLine(JsonSerializer.Serialize(new
        {
            output = outputPath,
            validation_errors = validationErrors.Count,
            first_errors = validationErrors
        }, new JsonSerializerOptions { WriteIndented = true }));

        return 0;
    }

    private static void RenderPlan(Body body, CreateDocxPlan plan)
    {
        body.RemoveAllChildren<Paragraph>();

        body.Append(CreateParagraph(plan.Title, centered: true, bold: true, fontSizeHalfPoints: 32));

        foreach (var subtitle in plan.Subtitles)
        {
            body.Append(CreateParagraph(subtitle, centered: true, italic: true, fontSizeHalfPoints: 24));
        }

        if (!string.IsNullOrWhiteSpace(plan.AuthorLine))
        {
            body.Append(CreateParagraph(plan.AuthorLine, centered: true, fontSizeHalfPoints: 22));
        }

        foreach (var paragraph in plan.Paragraphs)
        {
            body.Append(CreateParagraph(paragraph, centered: false, fontSizeHalfPoints: 24));
        }

        foreach (var section in plan.Sections)
        {
            body.Append(CreateParagraph(section.Heading, centered: false, bold: true, fontSizeHalfPoints: section.Level <= 1 ? 24 : 22));
            foreach (var paragraph in section.Paragraphs)
            {
                body.Append(CreateParagraph(paragraph, centered: false, fontSizeHalfPoints: 24));
            }
        }

        if (plan.References.Count > 0)
        {
            body.Append(CreateParagraph("REFERÊNCIAS", centered: false, bold: true, fontSizeHalfPoints: 24));
            foreach (var reference in plan.References)
            {
                body.Append(CreateParagraph(reference, centered: false, fontSizeHalfPoints: 22));
            }
        }
    }

    private static Paragraph CreateParagraph(string text, bool centered, bool bold = false, bool italic = false, int fontSizeHalfPoints = 24)
    {
        var paragraph = new Paragraph();
        var properties = new ParagraphProperties();
        properties.Append(new Justification { Val = centered ? JustificationValues.Center : JustificationValues.Both });
        paragraph.Append(properties);

        var run = new Run();
        var runProperties = new RunProperties();
        if (bold)
        {
            runProperties.Append(new Bold());
        }
        if (italic)
        {
            runProperties.Append(new Italic());
        }
        runProperties.Append(new FontSize { Val = fontSizeHalfPoints.ToString(System.Globalization.CultureInfo.InvariantCulture) });
        runProperties.Append(new FontSizeComplexScript { Val = fontSizeHalfPoints.ToString(System.Globalization.CultureInfo.InvariantCulture) });
        run.Append(runProperties);
        run.Append(new Text(text ?? "") { Space = SpaceProcessingModeValues.Preserve });
        paragraph.Append(run);
        return paragraph;
    }

    private static JsonSerializerOptions JsonOptions() => new()
    {
        PropertyNameCaseInsensitive = true
    };
}

internal sealed record CreateDocxPlan
{
    public string Title { get; init; } = "";
    public IReadOnlyList<string> Paragraphs { get; init; } = [];
    public IReadOnlyList<string> Subtitles { get; init; } = [];
    public string? AuthorLine { get; init; }
    public IReadOnlyList<CreateDocxSection> Sections { get; init; } = [];
    public IReadOnlyList<string> References { get; init; } = [];
}

internal sealed record CreateDocxSection
{
    public string Heading { get; init; } = "";
    public int Level { get; init; } = 1;
    public IReadOnlyList<string> Paragraphs { get; init; } = [];
}
