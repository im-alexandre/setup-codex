using System.Text.Json;
using DocumentFormat.OpenXml.Packaging;
using P = DocumentFormat.OpenXml.Presentation;
using A = DocumentFormat.OpenXml.Drawing;

namespace PptxOpenXmlTools.Commands;

internal static class InspectCommand
{
    public static int Run(string[] args)
    {
        if (args.Length < 2)
        {
            Console.Error.WriteLine("Usage: pptx-utils inspect <pptx> [--format json]");
            return 2;
        }

        if (!File.Exists(args[1]))
        {
            Console.Error.WriteLine($"File not found: {args[1]}");
            return 2;
        }

        if (!TryOpenPresentation(args[1], out var document))
        {
            return 2;
        }

        using (document)
        {
            if (document.PresentationPart?.Presentation?.SlideIdList is null)
            {
                Console.Error.WriteLine($"Invalid PPTX: missing presentation slide list: {args[1]}");
                return 2;
            }

            var slides = PresentationSlides.InPresentationOrder(document.PresentationPart);
            var result = new
            {
                slideCount = slides.Count,
                slides = slides.Select((slide, index) => new
                {
                    slideIndex = index + 1,
                    shapeCount = slide.Slide.Descendants<P.Shape>().Count(),
                    textShapeCount = slide.Slide.Descendants<P.Shape>().Count(s => s.Descendants<A.Text>().Any())
                })
            };

            Console.WriteLine(JsonSerializer.Serialize(result));
            return 0;
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
