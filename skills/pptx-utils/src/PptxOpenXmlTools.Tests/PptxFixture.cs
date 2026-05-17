using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using A = DocumentFormat.OpenXml.Drawing;

namespace PptxOpenXmlTools.Tests;

internal static class PptxFixture
{
    public static string CreateSingleSlideDeck(string text)
    {
        return CreateDeck([text], [0]);
    }

    public static string CreateDeckWithSlidePartInsertionOrderDifferentFromPresentationOrder()
    {
        return CreateDeck(["Second logical slide", "First logical slide"], [1, 0]);
    }

    public static string CreateSingleSlideDeckWithTextShapes(params string[] texts)
    {
        var path = Path.Combine(Path.GetTempPath(), $"pptx-utils-{Guid.NewGuid():N}.pptx");

        using var document = PresentationDocument.Create(path, PresentationDocumentType.Presentation);
        var presentationPart = document.AddPresentationPart();
        presentationPart.Presentation = new Presentation();
        var masterRelId = AddSlideMaster(presentationPart);

        var slidePart = presentationPart.AddNewPart<SlidePart>();
        var shapeTree = new ShapeTree(
            new NonVisualGroupShapeProperties(
                new NonVisualDrawingProperties { Id = 1U, Name = "" },
                new NonVisualGroupShapeDrawingProperties(),
                new ApplicationNonVisualDrawingProperties()),
            new GroupShapeProperties(new A.TransformGroup()));

        for (var index = 0; index < texts.Length; index++)
        {
            shapeTree.Append(CreateTextShape((uint)(index + 2), $"TextBox {index + 1}", texts[index]));
        }

        slidePart.Slide = new Slide(new CommonSlideData(shapeTree));
        var relId = presentationPart.GetIdOfPart(slidePart);
        presentationPart.Presentation.Append(
            new SlideMasterIdList(new SlideMasterId { Id = 2147483648U, RelationshipId = masterRelId }),
            new SlideIdList(new SlideId { Id = 256U, RelationshipId = relId }),
            new SlideSize { Cx = 9144000, Cy = 5143500, Type = SlideSizeValues.Screen16x9 },
            new NotesSize { Cx = 6858000, Cy = 9144000 });
        presentationPart.Presentation.Save();
        return path;
    }

    private static string CreateDeck(string[] textsInPartInsertionOrder, int[] presentationOrder)
    {
        var path = Path.Combine(Path.GetTempPath(), $"pptx-utils-{Guid.NewGuid():N}.pptx");

        using var document = PresentationDocument.Create(path, PresentationDocumentType.Presentation);
        var presentationPart = document.AddPresentationPart();
        presentationPart.Presentation = new Presentation();
        var masterRelId = AddSlideMaster(presentationPart);

        var slideParts = new List<SlidePart>();
        for (var index = 0; index < textsInPartInsertionOrder.Length; index++)
        {
            var slidePart = presentationPart.AddNewPart<SlidePart>();
            slidePart.Slide = new Slide(new CommonSlideData(new ShapeTree(
                new NonVisualGroupShapeProperties(
                    new NonVisualDrawingProperties { Id = 1U, Name = "" },
                    new NonVisualGroupShapeDrawingProperties(),
                    new ApplicationNonVisualDrawingProperties()),
                new GroupShapeProperties(new A.TransformGroup()),
                CreateTextShape(2U, "TextBox 1", textsInPartInsertionOrder[index]))));
            slideParts.Add(slidePart);
        }

        var slideIdList = new SlideIdList();
        for (var index = 0; index < presentationOrder.Length; index++)
        {
            var slidePart = slideParts[presentationOrder[index]];
            slideIdList.Append(new SlideId
            {
                Id = (uint)(256 + index),
                RelationshipId = presentationPart.GetIdOfPart(slidePart)
            });
        }

        presentationPart.Presentation.Append(
            new SlideMasterIdList(new SlideMasterId { Id = 2147483648U, RelationshipId = masterRelId }),
            slideIdList,
            new SlideSize { Cx = 9144000, Cy = 5143500, Type = SlideSizeValues.Screen16x9 },
            new NotesSize { Cx = 6858000, Cy = 9144000 });
        presentationPart.Presentation.Save();
        return path;
    }

    private static string AddSlideMaster(PresentationPart presentationPart)
    {
        var slideMasterPart = presentationPart.AddNewPart<SlideMasterPart>();
        slideMasterPart.SlideMaster = new SlideMaster(
            new CommonSlideData(CreateShapeTree()),
            new ColorMap
            {
                Background1 = A.ColorSchemeIndexValues.Light1,
                Text1 = A.ColorSchemeIndexValues.Dark1,
                Background2 = A.ColorSchemeIndexValues.Light2,
                Text2 = A.ColorSchemeIndexValues.Dark2,
                Accent1 = A.ColorSchemeIndexValues.Accent1,
                Accent2 = A.ColorSchemeIndexValues.Accent2,
                Accent3 = A.ColorSchemeIndexValues.Accent3,
                Accent4 = A.ColorSchemeIndexValues.Accent4,
                Accent5 = A.ColorSchemeIndexValues.Accent5,
                Accent6 = A.ColorSchemeIndexValues.Accent6,
                Hyperlink = A.ColorSchemeIndexValues.Hyperlink,
                FollowedHyperlink = A.ColorSchemeIndexValues.FollowedHyperlink
            });
        slideMasterPart.SlideMaster.Save();
        return presentationPart.GetIdOfPart(slideMasterPart);
    }

    private static ShapeTree CreateShapeTree()
    {
        return new ShapeTree(
            new NonVisualGroupShapeProperties(
                new NonVisualDrawingProperties { Id = 1U, Name = "" },
                new NonVisualGroupShapeDrawingProperties(),
                new ApplicationNonVisualDrawingProperties()),
            new GroupShapeProperties(new A.TransformGroup()));
    }

    private static Shape CreateTextShape(uint id, string name, string text)
    {
        return new Shape(
            new NonVisualShapeProperties(
                new NonVisualDrawingProperties { Id = id, Name = name },
                new NonVisualShapeDrawingProperties(new A.ShapeLocks { NoGrouping = true }),
                new ApplicationNonVisualDrawingProperties()),
            new ShapeProperties(),
            new TextBody(
                new A.BodyProperties(),
                new A.ListStyle(),
                new A.Paragraph(new A.Run(new A.Text(text)))));
    }
}
