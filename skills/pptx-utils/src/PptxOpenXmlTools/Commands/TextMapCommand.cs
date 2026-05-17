using System.Text.Json;
using DocumentFormat.OpenXml.Packaging;
using P = DocumentFormat.OpenXml.Presentation;
using A = DocumentFormat.OpenXml.Drawing;

namespace PptxOpenXmlTools.Commands;

internal static class TextMapCommand
{
    public static int Run(string[] args)
    {
        if (args.Length < 2)
        {
            Console.Error.WriteLine("Usage: pptx-utils text-map <pptx> [--format json]");
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

            var items = new List<object>();
            var slideIndex = 0;

            foreach (var slide in PresentationSlides.InPresentationOrder(document.PresentationPart))
            {
                slideIndex++;
                var shapeIndex = 0;

                foreach (var shape in slide.Slide.Descendants<P.Shape>())
                {
                    var text = string.Concat(shape.Descendants<A.Text>().Select(t => t.Text));
                    if (string.IsNullOrWhiteSpace(text))
                    {
                        continue;
                    }

                    shapeIndex++;
                    var name = shape.NonVisualShapeProperties?.NonVisualDrawingProperties?.Name?.Value ?? "";
                    items.Add(new
                    {
                        slideIndex,
                        shapeIndex,
                        shapeName = name,
                        text,
                        paragraphCount = shape.Descendants<A.Paragraph>().Count(),
                        runCount = shape.Descendants<A.Run>().Count()
                    });
                }
            }

            Console.WriteLine(JsonSerializer.Serialize(new { items }));
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
