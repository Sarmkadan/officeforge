using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using OfficeForge.Models;
using A = DocumentFormat.OpenXml.Drawing;

namespace OfficeForge.Slides;

/// <summary>
/// Writes a PowerPoint presentation from a <see cref="PresentationModel"/> using the OpenXML SDK.
/// </summary>
public sealed class PptxWriter : IDocumentWriter<PresentationModel>
{
    /// <summary>
    /// Writes a PowerPoint presentation to a file.
    /// </summary>
    /// <param name="model">The presentation model to write.</param>
    /// <param name="path">The file path to write to.</param>
    /// <exception cref="ArgumentNullException"><paramref name="model"/> is <see langword="null"/></exception>
    /// <exception cref="ArgumentException"><paramref name="path"/> is <see langword="null"/>, empty, or whitespace</exception>
    public void Write(PresentationModel model, string path)
    {
        ArgumentNullException.ThrowIfNull(model);
        ArgumentException.ThrowIfNullOrEmpty(path);

        using var stream = File.Create(path);
        Write(model, stream);
    }

    /// <summary>
    /// Writes a PowerPoint presentation to a stream.
    /// </summary>
    /// <param name="model">The presentation model to write.</param>
    /// <param name="stream">The stream to write to.</param>
    /// <exception cref="ArgumentNullException"><paramref name="model"/> or <paramref name="stream"/> is <see langword="null"/></exception>
    public void Write(PresentationModel model, Stream stream)
    {
        ArgumentNullException.ThrowIfNull(model);
        ArgumentNullException.ThrowIfNull(stream);

        using var document = PresentationDocument.Create(stream, DocumentFormat.OpenXml.PresentationDocumentType.Presentation);
        var presentationPart = document.AddPresentationPart();
        presentationPart.Presentation = new Presentation();

        var slideIdList = presentationPart.Presentation.SlideIdList ??= new SlideIdList();
        var slideId = 1u;

        foreach (var slideModel in model.Slides)
        {
            var slidePart = presentationPart.AddNewPart<SlidePart>();
            slidePart.Slide = CreateSlide(slideModel);

            slideIdList.Append(new SlideId
            {
                Id = slideId,
                RelationshipId = presentationPart.GetIdOfPart(slidePart)
            });

            slideId++;
        }

        presentationPart.Presentation.Save();
    }

    private static Slide CreateSlide(SlideModel slideModel)
    {
        var commonSlideData = new CommonSlideData();
        var shapeTree = new ShapeTree();

        if (slideModel.Title is not null)
        {
            shapeTree.Append(CreateTitlePlaceholder(slideModel.Title));
        }

        foreach (var shapeModel in slideModel.Shapes)
        {
            shapeTree.Append(CreateContentShape(shapeModel));
        }

        commonSlideData.Append(shapeTree);
        return new Slide(commonSlideData, new ColorMapOverride(new A.MasterColorMapping()));
    }

    private static Shape CreateTitlePlaceholder(string title)
    {
        return new Shape(
            new NonVisualShapeProperties(
                new NonVisualDrawingProperties { Name = "Title" },
                new NonVisualShapeDrawingProperties(),
                new ApplicationNonVisualDrawingProperties(
                    new PlaceholderShape { Type = PlaceholderValues.Title }
                )
            ),
            new ShapeProperties(),
            new TextBody(
                new A.BodyProperties(),
                new A.ListStyle(),
                new A.Paragraph(
                    new A.Run(
                        new A.RunProperties(),
                        new A.Text(title)
                    )
                )
            )
        );
    }

    private static Shape CreateContentShape(ShapeTextModel shapeModel)
    {
        var text = string.Join("\n", shapeModel.Lines);

        return new Shape(
            new NonVisualShapeProperties(
                new NonVisualDrawingProperties { Name = shapeModel.Name ?? "Content Placeholder" },
                new NonVisualShapeDrawingProperties(),
                new ApplicationNonVisualDrawingProperties()
            ),
            new ShapeProperties(),
            new TextBody(
                new A.BodyProperties(),
                new A.ListStyle(),
                new A.Paragraph(
                    new A.Run(
                        new A.RunProperties(),
                        new A.Text(text)
                    )
                )
            )
        );
    }
}
