using System.Globalization;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeForge.Models;

namespace OfficeForge.Words;

/// <summary>
/// Writes a <see cref="DocumentModel"/> to a DOCX stream or file.
/// </summary>
public sealed class DocxWriter : IDocumentWriter<DocumentModel>
{
    /// <summary>
    /// Writes the supplied <paramref name="model"/> to the given <paramref name="stream"/>.
    /// </summary>
    /// <param name="model">The document model to write.</param>
    /// <param name="stream">The target stream.</param>
    /// <exception cref="ArgumentNullException">Thrown when <paramref name="model"/> or <paramref name="stream"/> is <c>null</c>.</exception>
    public void Write(DocumentModel model, Stream stream)
    {
        ArgumentNullException.ThrowIfNull(model);
        ArgumentNullException.ThrowIfNull(stream);
        using var document = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document);
        var mainPart = document.AddMainDocumentPart();
        var body = new Body();
        foreach (var paragraphModel in model.Paragraphs)
        {
            var paragraph = new Paragraph();
            if (paragraphModel.Kind == ParagraphKind.Heading && paragraphModel.HeadingLevel is > 0 and <= 9)
                paragraph.ParagraphProperties = new ParagraphProperties(
                    new ParagraphStyleId { Val = $"Heading{paragraphModel.HeadingLevel}" });

            foreach (var runModel in paragraphModel.Runs)
            {
                var run = new Run();
                var props = new RunProperties();

                if (runModel.Style.Bold) props.Append(new Bold());
                if (runModel.Style.Italic) props.Append(new Italic());
                if (runModel.Style.Underline) props.Append(new Underline { Val = UnderlineValues.Single });
                if (runModel.Style.FontName is { } font) props.Append(new RunFonts { Ascii = font });
                if (runModel.Style.FontSize is double size)
                    props.Append(new FontSize { Val = size.ToString(CultureInfo.InvariantCulture) });

                if (props.HasChildren) run.RunProperties = props;
                run.Append(new Text(runModel.Text) { Space = SpaceProcessingModeValues.Preserve });
                paragraph.Append(run);
            }

            body.Append(paragraph);
        }

        mainPart.Document = new Document(body);
        mainPart.Document.Save();
    }

    /// <summary>
    /// Asynchronously writes the supplied <paramref name="model"/> to the given <paramref name="stream"/>.
    /// </summary>
    /// <param name="model">The document model to write.</param>
    /// <param name="stream">The target stream.</param>
    /// <param name="cancellationToken">A token to monitor for cancellation requests.</param>
    /// <returns>A task that represents the asynchronous write operation.</returns>
    /// <exception cref="ArgumentNullException">Thrown when <paramref name="model"/> or <paramref name="stream"/> is <c>null</c>.</exception>
    public async Task WriteAsync(DocumentModel model, Stream stream, CancellationToken cancellationToken = default)
    {
        ArgumentNullException.ThrowIfNull(model);
        ArgumentNullException.ThrowIfNull(stream);

        // Write to a temporary memory stream synchronously, then copy asynchronously to the target stream.
        using var tempStream = new MemoryStream();
        Write(model, tempStream);
        tempStream.Position = 0;
        await tempStream.CopyToAsync(stream, 81920, cancellationToken).ConfigureAwait(false);
    }

    /// <summary>
    /// Writes the supplied <paramref name="model"/> to a file at <paramref name="path"/>.
    /// </summary>
    /// <param name="model">The document model to write.</param>
    /// <param name="path">The destination file path.</param>
    /// <exception cref="ArgumentNullException">Thrown when <paramref name="model"/> or <paramref name="path"/> is <c>null</c>.</exception>
    /// <exception cref="ArgumentException">Thrown when <paramref name="path"/> is an empty string.</exception>
    public void Write(DocumentModel model, string path)
    {
        ArgumentNullException.ThrowIfNull(model);
        ArgumentNullException.ThrowIfNull(path);
        ArgumentException.ThrowIfNullOrEmpty(path);
        using var stream = File.Create(path);
        Write(model, stream);
    }

    /// <summary>
    /// Asynchronously writes the supplied <paramref name="model"/> to a file at <paramref name="path"/>.
    /// </summary>
    /// <param name="model">The document model to write.</param>
    /// <param name="path">The destination file path.</param>
    /// <param name="cancellationToken">A token to monitor for cancellation requests.</param>
    /// <returns>A task that represents the asynchronous write operation.</returns>
    /// <exception cref="ArgumentNullException">Thrown when <paramref name="model"/> or <paramref name="path"/> is <c>null</c>.</exception>
    /// <exception cref="ArgumentException">Thrown when <paramref name="path"/> is an empty string.</exception>
    public async Task WriteAsync(DocumentModel model, string path, CancellationToken cancellationToken = default)
    {
        ArgumentNullException.ThrowIfNull(model);
        ArgumentNullException.ThrowIfNull(path);
        ArgumentException.ThrowIfNullOrEmpty(path);

        // Open the file stream with async enabled and delegate to the stream overload.
        await using var stream = new FileStream(
            path,
            FileMode.Create,
            FileAccess.Write,
            FileShare.None,
            bufferSize: 4096,
            useAsync: true);

        await WriteAsync(model, stream, cancellationToken).ConfigureAwait(false);
    }
}
