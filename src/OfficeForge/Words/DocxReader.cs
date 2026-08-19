using System.IO;
using System.IO.Compression;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeForge.Models;

namespace OfficeForge.Words;

/// <summary>
/// Reads a Word document (<c>.docx</c>) and produces a <see cref="DocumentModel"/>.
/// </summary>
public sealed class DocxReader : IDocumentReader<DocumentModel>, IEquatable<DocxReader>
{
    private readonly ReaderOptions _defaultOptions;

    /// <summary>
    /// Initializes a new instance of <see cref="DocxReader"/> using the default <see cref="ReaderOptions"/>.
    /// </summary>
    public DocxReader() : this(ReaderOptions.Default) { }

    /// <summary>
    /// Initializes a new instance of <see cref="DocxReader"/> with the specified <paramref name="options"/>.
    /// </summary>
    /// <param name="options">The options that control size and entry limits.</param>
    /// <exception cref="ArgumentNullException">Thrown when <paramref name="options"/> is <c>null</c>.</exception>
    public DocxReader(ReaderOptions options)
    {
        ArgumentNullException.ThrowIfNull(options);
        _defaultOptions = options;
    }

    /// <inheritdoc />
    public DocumentModel Read(Stream stream) => Read(stream, _defaultOptions);

    /// <inheritdoc />
    public DocumentModel Read(Stream stream, ReaderOptions options)
    {
        ArgumentNullException.ThrowIfNull(stream);
        ArgumentNullException.ThrowIfNull(options);
        options.Validate();

        // Check raw stream length before opening as zip (pre‑decompression size)
        if (stream.CanSeek && options.MaxUncompressedSize > 0 && stream.Length > options.MaxUncompressedSize)
        {
            throw new DocumentTooLargeException(
                $"Document exceeds maximum uncompressed size limit of {options.MaxUncompressedSize} bytes (actual: {stream.Length} bytes).")
            {
                LimitType = DocumentTooLargeException.SizeLimitType.MaxUncompressedSize,
                MaxLimit = options.MaxUncompressedSize,
                ActualValue = stream.Length
            };
        }

        // Open as zip to enforce entry‑count, compression‑ratio and total decompressed size limits
        if (stream.CanSeek)
        {
            stream.Position = 0;
            using var zipArchive = new ZipArchive(stream, ZipArchiveMode.Read, false);

            // Entry count limit
            if (options.MaxEntryCount > 0 && zipArchive.Entries.Count > options.MaxEntryCount)
            {
                throw new DocumentTooLargeException(
                    $"Document contains {zipArchive.Entries.Count} entries, exceeding maximum entry count limit of {options.MaxEntryCount}.")
                {
                    LimitType = DocumentTooLargeException.SizeLimitType.MaxEntryCount,
                    MaxLimit = options.MaxEntryCount,
                    ActualValue = zipArchive.Entries.Count
                };
            }

            // Total decompressed size limit
            if (options.MaxUncompressedSize > 0)
            {
                long totalDecompressedSize = zipArchive.Entries.Sum(e => e.Length);
                if (totalDecompressedSize > options.MaxUncompressedSize)
                {
                    throw new DocumentTooLargeException(
                        $"Document total decompressed size {totalDecompressedSize} bytes exceeds limit of {options.MaxUncompressedSize} bytes.")
                    {
                        LimitType = DocumentTooLargeException.SizeLimitType.MaxUncompressedSize,
                        MaxLimit = options.MaxUncompressedSize,
                        ActualValue = totalDecompressedSize
                    };
                }
            }

            // Compression ratio limit (sampled)
            if (options.MaxCompressionRatio > 0)
            {
                long totalCompressedSize = 0;
                long totalUncompressedSize = 0;
                int entriesSampled = 0;
                const int maxEntriesToSample = 100;

                foreach (var entry in zipArchive.Entries)
                {
                    if (entriesSampled >= maxEntriesToSample) break;
                    if (entry.Length <= 0) continue;

                    totalCompressedSize += entry.CompressedLength;
                    totalUncompressedSize += entry.Length;
                    entriesSampled++;
                }

                if (entriesSampled > 0 && totalCompressedSize > 0)
                {
                    double ratio = (double)totalUncompressedSize / totalCompressedSize;
                    if (ratio > options.MaxCompressionRatio)
                    {
                        throw new DocumentTooLargeException(
                            $"Document compression ratio {ratio:N2}:1 exceeds maximum allowed ratio of {options.MaxCompressionRatio}:1.")
                        {
                            LimitType = DocumentTooLargeException.SizeLimitType.MaxCompressionRatio,
                            MaxLimit = options.MaxCompressionRatio,
                            ActualValue = (long)ratio
                        };
                    }
                }
            }

            stream.Position = 0;
        }

        WordprocessingDocument document;
        try
        {
            document = WordprocessingDocument.Open(stream, false);
        }
        catch (OpenXmlPackageException ex)
        {
            throw new OfficeForgeFormatException("The document is not a valid Word document.", ex);
        }

        using (document)
        {
            var body = document.MainDocumentPart?.Document.Body ?? throw new InvalidDataException("Document body missing.");

            // Document part size check
            if (document.MainDocumentPart != null)
            {
                try
                {
                    using var partStream = document.MainDocumentPart.GetStream();
                    var partLength = partStream.Length;
                    if (options.MaxPartSize > 0 && partLength > options.MaxPartSize)
                    {
                        throw new DocumentTooLargeException(
                            $"Document part exceeds maximum size limit of {options.MaxPartSize} bytes (actual: {partLength} bytes).")
                        {
                            LimitType = DocumentTooLargeException.SizeLimitType.MaxPartSize,
                            MaxLimit = options.MaxPartSize,
                            ActualValue = partLength,
                            PartName = "Document"
                        };
                    }
                }
                catch (IOException)
                {
                    // Stream is not accessible, skip size check
                }
            }

            var model = new DocumentModel();
            foreach (var paragraph in body.Descendants<Paragraph>())
            {
                var paragraphModel = new ParagraphModel();
                var styleId = paragraph.ParagraphProperties?.ParagraphStyleId?.Val?.Value;
                if (styleId is not null && styleId.StartsWith("Heading", StringComparison.OrdinalIgnoreCase))
                {
                    paragraphModel.Kind = ParagraphKind.Heading;
                    if (int.TryParse(styleId.AsSpan("Heading".Length), out var level))
                        paragraphModel.HeadingLevel = level;
                }

                foreach (var run in paragraph.Descendants<Run>())
                {
                    var text = string.Concat(run.Descendants<Text>().Select(t => t.Text));
                    if (text.Length == 0) continue;

                    var props = run.RunProperties;
                    var style = new Models.RunStyle(
                        Bold: IsOn(props?.Bold),
                        Italic: IsOn(props?.Italic),
                        Underline: props?.Underline is { } u && u.Val?.Value != UnderlineValues.None,
                        FontName: props?.RunFonts?.Ascii?.Value);

                    paragraphModel.Runs.Add(new RunModel(text, style));
                }

                model.Paragraphs.Add(paragraphModel);
            }

            return model;
        }
    }

    /// <inheritdoc />
    public DocumentModel Read(string path)
    {
        ArgumentNullException.ThrowIfNull(path);
        ArgumentException.ThrowIfNullOrEmpty(path);
        using var stream = File.OpenRead(path);
        return Read(stream);
    }

    /// <inheritdoc />
    public DocumentModel Read(string path, ReaderOptions options)
    {
        ArgumentNullException.ThrowIfNull(path);
        ArgumentException.ThrowIfNullOrEmpty(path);
        ArgumentNullException.ThrowIfNull(options);
        options.Validate();

        using var stream = File.OpenRead(path);
        return Read(stream, options);
    }

    // OnOff semantics: the bare element (<w:b/>) means on; an explicit
    // w:val ("0"/"false") can turn it off.
    private static bool IsOn(OnOffType? property) =>
        property is not null && (property.Val is null || property.Val.Value);

    // IEquatable<DocxReader> implementation
    public bool Equals(DocxReader? other)
    {
        if (other is null) return false;
        return _defaultOptions.Equals(other._defaultOptions);
    }

    public override bool Equals(object? obj) => Equals(obj as DocxReader);

    public override int GetHashCode() => _defaultOptions?.GetHashCode() ?? 0;

    public static bool operator ==(DocxReader? left, DocxReader? right) =>
        EqualityComparer<DocxReader>.Default.Equals(left, right);

    public static bool operator !=(DocxReader? left, DocxReader? right) =>
        !(left == right);
}