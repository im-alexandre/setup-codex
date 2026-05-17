using System.Text.Json;
using DocumentFormat.OpenXml.Packaging;
using PptxOpenXmlTools.Models;
using P = DocumentFormat.OpenXml.Presentation;
using A = DocumentFormat.OpenXml.Drawing;

namespace PptxOpenXmlTools.Commands;

internal static class ReplaceTextCommand
{
    public static int Run(string[] args)
    {
        if (args.Length < 6)
        {
            Console.Error.WriteLine("Usage: pptx-utils replace-text <pptx> --plan <json> --output <pptx>");
            return 2;
        }

        var input = args[1];
        var planPath = ValueAfter(args, "--plan");
        var output = ValueAfter(args, "--output");
        if (planPath is null || output is null)
        {
            Console.Error.WriteLine("Missing --plan or --output.");
            return 2;
        }

        if (!File.Exists(input))
        {
            Console.Error.WriteLine($"File not found: {input}");
            return 2;
        }

        if (!File.Exists(planPath))
        {
            Console.Error.WriteLine($"Plan not found: {planPath}");
            return 2;
        }

        ReplaceTextPlan? plan;
        try
        {
            plan = JsonSerializer.Deserialize<ReplaceTextPlan>(
                File.ReadAllText(planPath),
                new JsonSerializerOptions { PropertyNameCaseInsensitive = true });
        }
        catch (JsonException)
        {
            Console.Error.WriteLine("Invalid replace-text plan.");
            return 2;
        }

        if (plan is null)
        {
            Console.Error.WriteLine("Invalid replace-text plan.");
            return 2;
        }

        if (plan.Replacements is null || plan.Replacements.Count == 0)
        {
            Console.Error.WriteLine("No replacements in replace-text plan.");
            return 2;
        }

        foreach (var replacement in plan.Replacements)
        {
            var validationError = ValidateReplacement(replacement);
            if (validationError is not null)
            {
                Console.Error.WriteLine(validationError);
                return 2;
            }
        }

        var outputDirectory = Path.GetDirectoryName(output);
        if (!string.IsNullOrWhiteSpace(outputDirectory) && !Directory.Exists(outputDirectory))
        {
            Console.Error.WriteLine($"Output directory not found: {outputDirectory}");
            return 2;
        }

        if (!TryOpenPresentation(input, false, out var inputDocument))
        {
            return 2;
        }

        using (inputDocument)
        {
            var inputSlides = PresentationSlides.InPresentationOrder(inputDocument.PresentationPart!);
            foreach (var replacement in plan.Replacements)
            {
                var validationExitCode = ValidateReplacementAgainstSlides(replacement, inputSlides);
                if (validationExitCode != 0)
                {
                    return validationExitCode;
                }
            }
        }

        try
        {
            File.Copy(input, output, overwrite: true);
        }
        catch (IOException)
        {
            Console.Error.WriteLine($"Unable to write output: {output}");
            return 2;
        }
        catch (UnauthorizedAccessException)
        {
            Console.Error.WriteLine($"Unable to write output: {output}");
            return 2;
        }

        if (!TryOpenPresentation(output, true, out var outputDocument))
        {
            TryDelete(output);
            return 2;
        }

        using (outputDocument)
        {
            var outputSlides = PresentationSlides.InPresentationOrder(outputDocument.PresentationPart!);

            foreach (var replacement in plan.Replacements)
            {
                var slide = outputSlides[replacement.SlideIndex!.Value - 1];
                var shapes = slide.Slide.Descendants<P.Shape>().Where(s => s.Descendants<A.Text>().Any()).ToList();
                var shape = shapes[replacement.ShapeIndex!.Value - 1];
                var textNodes = shape.Descendants<A.Text>().ToList();

                textNodes[0].Text = replacement.Replace!;
                foreach (var extra in textNodes.Skip(1))
                {
                    extra.Text = "";
                }

                slide.Slide.Save();
            }
        }

        return 0;
    }

    private static string? ValueAfter(string[] args, string option)
    {
        var index = Array.IndexOf(args, option);
        return index >= 0 && index + 1 < args.Length ? args[index + 1] : null;
    }

    private static void TryDelete(string path)
    {
        try
        {
            if (File.Exists(path))
            {
                File.Delete(path);
            }
        }
        catch (IOException)
        {
        }
        catch (UnauthorizedAccessException)
        {
        }
    }

    private static string? ValidateReplacement(ReplaceTextItem replacement)
    {
        if (replacement.SlideIndex is null)
        {
            return "Missing slideIndex in replace-text plan.";
        }

        if (replacement.ShapeIndex is null)
        {
            return "Missing shapeIndex in replace-text plan.";
        }

        if (string.IsNullOrWhiteSpace(replacement.Find))
        {
            return "Missing find in replace-text plan.";
        }

        if (replacement.Replace is null)
        {
            return "Missing replace in replace-text plan.";
        }

        if (string.IsNullOrWhiteSpace(replacement.Mode))
        {
            return "Missing mode in replace-text plan.";
        }

        return null;
    }

    private static int ValidateReplacementAgainstSlides(ReplaceTextItem replacement, IReadOnlyList<SlidePart> slides)
    {
        if (replacement.Mode != "exact")
        {
            Console.Error.WriteLine($"Unsupported mode: {replacement.Mode}");
            return 2;
        }

        var slideIndex = replacement.SlideIndex!.Value;
        var shapeIndex = replacement.ShapeIndex!.Value;

        if (slideIndex < 1 || slideIndex > slides.Count)
        {
            Console.Error.WriteLine($"Slide index out of range: {slideIndex}");
            return 2;
        }

        var slide = slides[slideIndex - 1];
        var shapes = slide.Slide.Descendants<P.Shape>().Where(s => s.Descendants<A.Text>().Any()).ToList();
        if (shapeIndex < 1 || shapeIndex > shapes.Count)
        {
            Console.Error.WriteLine($"Shape index out of range: {shapeIndex}");
            return 2;
        }

        var shape = shapes[shapeIndex - 1];
        var current = string.Concat(shape.Descendants<A.Text>().Select(t => t.Text));
        if (current != replacement.Find)
        {
            Console.Error.WriteLine($"Text mismatch on slide {slideIndex}, shape {shapeIndex}.");
            return 1;
        }

        return 0;
    }

    private static bool TryOpenPresentation(string pptx, bool isEditable, out PresentationDocument? document)
    {
        try
        {
            document = PresentationDocument.Open(pptx, isEditable);
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
