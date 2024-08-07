using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using OfficeForge.Models;
using A = DocumentFormat.OpenXml.Drawing;

namespace OfficeForge.Slides;

public sealed class PptxReader : IDocumentReader<PresentationModel>
{
    public PresentationModel Read(string path)
    {
        using var stream = File.OpenRead(path);
        return Read(stream);
    }

    public PresentationModel Read(Stream stream)
    {
        using var document = PresentationDocument.Open(stream, false);
        var presentationPart = document.PresentationPart ?? throw new InvalidDataException("Presentation part missing.");
        var model = new PresentationModel();
        var slideIds = presentationPart.Presentation.SlideIdList?.Elements<SlideId>() ?? [];
        foreach (var slideId in slideIds)
        {
            if (slideId.RelationshipId?.Value is not { } relId) continue;
            if (presentationPart.GetPartById(relId) is not SlidePart slidePart) continue;
            var slide = model.AddSlide();
            foreach (var shape in slidePart.Slide.Descendants<Shape>())
            {
                var lines = shape.TextBody?.Descendants<A.Paragraph>()
                    .Select(p => string.Concat(p.Descendants<A.Text>().Select(t => t.Text)))
                    .Where(l => l.Length > 0)
                    .ToList() ?? [];
                if (lines.Count == 0) continue;
                var placeholder = shape.NonVisualShapeProperties?.ApplicationNonVisualDrawingProperties?
                    .PlaceholderShape?.Type?.Value;
                if (slide.Title is null && (placeholder == PlaceholderValues.Title || placeholder == PlaceholderValues.CenteredTitle))
                {
                    slide.Title = string.Join(" ", lines);
                    continue;
                }
                var shapeText = new ShapeTextModel { Name = shape.NonVisualShapeProperties?.NonVisualDrawingProperties?.Name?.Value };
                shapeText.Lines.AddRange(lines);
                slide.Shapes.Add(shapeText);
            }
        }
        return model;
    }
}
