using DocumentFormat.OpenXml.Packaging;
using P = DocumentFormat.OpenXml.Presentation;

namespace PptxOpenXmlTools.Commands;

internal static class PresentationSlides
{
    public static IReadOnlyList<SlidePart> InPresentationOrder(PresentationPart presentationPart)
    {
        var slideIds = presentationPart.Presentation.SlideIdList?.Elements<P.SlideId>()
            ?? Enumerable.Empty<P.SlideId>();

        return slideIds
            .Select(slideId => slideId.RelationshipId?.Value)
            .Where(relationshipId => !string.IsNullOrWhiteSpace(relationshipId))
            .Select(relationshipId => presentationPart.GetPartById(relationshipId!))
            .OfType<SlidePart>()
            .ToList();
    }
}
