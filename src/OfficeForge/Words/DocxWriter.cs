using System;
using System.Globalization;
using System.IO;
using System.IO.Compression;
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
    private const int TwipsPerPoint = 20;
    private const uint HeaderFooterTwips = 720;

    /// <summary>
    /// Gets the options that control default fonts, page margins and compression.
    /// </summary>
    public DocxWriterOptions Options { get; }

    /// <summary>
    /// Initializes a new instance of <see cref="DocxWriter"/> using <see cref="DocxWriterOptions.Default"/>.
    /// </summary>
    public DocxWriter() : this(DocxWriterOptions.Default) { }

    /// <summary>
    /// Initializes a new instance of <see cref="DocxWriter"/> with the specified <paramref name="options"/>.
    /// </summary>
    /// <param name="options">The writer options.</param>
    /// <exception cref="ArgumentNullException">Thrown when <paramref name="options"/> is <c>null</c>.</exception>
    /// <exception cref="ArgumentException">Thrown when <paramref name="options"/> fails validation.</exception>
    public DocxWriter(DocxWriterOptions options)
    {
        ArgumentNullException.ThrowIfNull(options);
        options.Validate();
        Options = options;
    }

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

        // Build the package in memory first so it can be re-compressed at the configured level.
        using var buffer = new MemoryStream();
        BuildPackage(model, buffer);
        Recompress(buffer, stream, Options.CompressionLevel);
    }

    private void BuildPackage(DocumentModel model, Stream stream)
    {
        using var document = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document);
        var mainPart = document.AddMainDocumentPart();

        var stylesPart = mainPart.AddNewPart<StyleDefinitionsPart>();
        var halfPoints = ToHalfPoints(Options.DefaultFontSize);
        stylesPart.Styles = new Styles(
            new DocDefaults(
                new RunPropertiesDefault(
                    new RunPropertiesBaseStyle(
                        new RunFonts { Ascii = Options.DefaultFontName, HighAnsi = Options.DefaultFontName },
                        new FontSize { Val = halfPoints },
                        new FontSizeComplexScript { Val = halfPoints }))));

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

        var margins = Options.PageMargins;
        body.Append(new SectionProperties(new PageMargin
        {
            Top = ToTwips(margins.Top),
            Bottom = ToTwips(margins.Bottom),
            Left = (uint)ToTwips(margins.Left),
            Right = (uint)ToTwips(margins.Right),
            Header = HeaderFooterTwips,
            Footer = HeaderFooterTwips,
            Gutter = 0
        }));

        mainPart.Document = new Document(body);
        mainPart.Document.Save();
    }

    /// <summary>
    /// Copies every entry of the package in <paramref name="source"/> into <paramref name="target"/>,
    /// compressing each entry with <paramref name="level"/>.
    /// </summary>
    private static void Recompress(Stream source, Stream target, CompressionLevel level)
    {
        source.Position = 0;
        using var input = new ZipArchive(source, ZipArchiveMode.Read, leaveOpen: true);
        using var output = new ZipArchive(target, ZipArchiveMode.Create, leaveOpen: true);
        foreach (var entry in input.Entries)
        {
            var copy = output.CreateEntry(entry.FullName, level);
            using var from = entry.Open();
            using var to = copy.Open();
            from.CopyTo(to);
        }
    }

    private static string ToHalfPoints(double points) =>
        ((int)Math.Round(points * 2, MidpointRounding.AwayFromZero)).ToString(CultureInfo.InvariantCulture);

    private static int ToTwips(double points) =>
        (int)Math.Round(points * TwipsPerPoint, MidpointRounding.AwayFromZero);

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
    /// <exception cref="ArgumentException">Thrown when <paramref name="path"/> is an empty string or contains invalid characters.</exception>
    public void Write(DocumentModel model, string path)
    {
        ArgumentNullException.ThrowIfNull(model);
        ArgumentNullException.ThrowIfNull(path);
        ArgumentException.ThrowIfNullOrEmpty(path);
        ValidatePath(path);
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
    /// <exception cref="ArgumentException">Thrown when <paramref name="path"/> is an empty string or contains invalid characters.</exception>
    public async Task WriteAsync(DocumentModel model, string path, CancellationToken cancellationToken = default)
    {
        ArgumentNullException.ThrowIfNull(model);
        ArgumentNullException.ThrowIfNull(path);
        ArgumentException.ThrowIfNullOrEmpty(path);
        ValidatePath(path);

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
}
