using System.IO;
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
    /// Gets the options that control slide size, the theme name, the slide limit and speaker notes.
    /// </summary>
    public PptxWriterOptions Options { get; }

    /// <summary>
    /// Initializes a new instance of <see cref="PptxWriter"/> using <see cref="PptxWriterOptions.Default"/>.
    /// </summary>
    public PptxWriter() : this(PptxWriterOptions.Default) { }

    /// <summary>
    /// Initializes a new instance of <see cref="PptxWriter"/> with the specified <paramref name="options"/>.
    /// </summary>
    /// <param name="options">The writer options.</param>
    /// <exception cref="ArgumentNullException">Thrown when <paramref name="options"/> is <see langword="null"/>.</exception>
    /// <exception cref="ArgumentException">Thrown when <paramref name="options"/> fails validation.</exception>
    public PptxWriter(PptxWriterOptions options)
    {
        ArgumentNullException.ThrowIfNull(options);
        options.Validate();
        Options = options;
    }

    /// <summary>
    /// Writes a PowerPoint presentation to a stream.
    /// </summary>
    /// <param name="model">The presentation model to write.</param>
    /// <param name="stream">The stream to write to.</param>
    /// <exception cref="ArgumentNullException"><paramref name="model"/> or <paramref name="stream"/> is <see langword="null"/></exception>
    /// <exception cref="InvalidOperationException"><paramref name="model"/> has more slides than <see cref="PptxWriterOptions.MaxSlides"/> allows.</exception>
    public void Write(PresentationModel model, Stream stream)
    {
        ArgumentNullException.ThrowIfNull(model);
        ArgumentNullException.ThrowIfNull(stream);
        EnsureWithinSlideLimit(model);

        using var document = PresentationDocument.Create(stream, DocumentFormat.OpenXml.PresentationDocumentType.Presentation);
        var presentationPart = document.AddPresentationPart();
        presentationPart.Presentation = new Presentation();
        presentationPart.Presentation.SlideSize = new SlideSize
        {
            Cx = Options.SlideWidth,
            Cy = Options.SlideHeight
        };

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

    /// <summary>
    /// Writes a PowerPoint presentation to a file.
    /// </summary>
    /// <param name="model">The presentation model to write.</param>
    /// <param name="path">The file path to write to.</param>
    /// <exception cref="ArgumentNullException"><paramref name="model"/> is <see langword="null"/></exception>
    /// <exception cref="ArgumentException"><paramref name="path"/> is <see langword="null"/>, empty, whitespace, or contains invalid characters</exception>
    /// <exception cref="InvalidOperationException"><paramref name="model"/> has more slides than <see cref="PptxWriterOptions.MaxSlides"/> allows.</exception>
    public void Write(PresentationModel model, string path)
    {
        ArgumentNullException.ThrowIfNull(model);
        ArgumentNullException.ThrowIfNull(path);
        ArgumentException.ThrowIfNullOrEmpty(path);
        ValidatePath(path);
        EnsureWithinSlideLimit(model);
        using var stream = File.Create(path);
        Write(model, stream);
    }

    private void EnsureWithinSlideLimit(PresentationModel model)
    {
        if (model.Slides.Count > Options.MaxSlides)
        {
            throw new InvalidOperationException(
                $"The presentation has {model.Slides.Count} slides, which exceeds {nameof(PptxWriterOptions.MaxSlides)} = {Options.MaxSlides}.");
        }
    }

    private static void ValidatePath(string path)
    {
        if (path.IndexOfAny(Path.GetInvalidPathChars()) >= 0)
        {
            throw new ArgumentException("The path contains invalid characters.", nameof(path));
        }

        // Check for path traversal attempts
        var normalizedPath = path.Replace(Path.AltDirectorySeparatorChar, Path.DirectorySeparatorChar);
        var parts = normalizedPath.Split(Path.DirectorySeparatorChar, StringSplitOptions.RemoveEmptyEntries);
        foreach (var part in parts)
        {
            if (part == ".." || part == ".")
            {
                throw new ArgumentException("The path cannot contain directory traversal sequences.", nameof(part));
            }
        }
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
