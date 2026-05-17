using System.Text.Json;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;

namespace PptxOpenXmlTools.Commands;

internal static class ValidateCommand
{
    public static int Run(string[] args)
    {
        if (args.Length < 2)
        {
            Console.Error.WriteLine("Usage: pptx-utils validate <pptx> [--format json]");
            return 2;
        }

        var pptx = args[1];
        if (!File.Exists(pptx))
        {
            Console.Error.WriteLine($"File not found: {pptx}");
            return 2;
        }

        if (!TryOpenPresentation(pptx, out var document))
        {
            return 2;
        }

        using (document)
        {
            var validator = new OpenXmlValidator();
            var errors = validator.Validate(document).ToList();
            var hasPresentation = document.PresentationPart?.Presentation is not null;
            var hasSlideIdList = document.PresentationPart?.Presentation?.SlideIdList is not null;

            var result = new
            {
                isValid = errors.Count == 0 && hasPresentation && hasSlideIdList,
                openXmlValidationErrors = errors.Count,
                hasPresentation,
                hasSlideIdList,
                errors = errors.Select(error => new
                {
                    description = error.Description,
                    path = error.Path?.XPath ?? ""
                }).ToArray()
            };

            Console.WriteLine(JsonSerializer.Serialize(result));
            return result.isValid ? 0 : 1;
        }
    }

    private static bool TryOpenPresentation(string pptx, out PresentationDocument? document)
    {
        try
        {
            document = PresentationDocument.Open(pptx, false);
            return true;
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"Invalid PPTX: {ex.Message}");
            document = null;
            return false;
        }
    }
}
