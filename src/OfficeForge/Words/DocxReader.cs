using System.IO;
using System.IO.Compression;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeForge.Models;

namespace OfficeForge.Words;

public sealed class DocxReader : IDocumentReader<DocumentModel>
{
    /// <inheritdoc />
    public DocumentModel Read(Stream stream)
    {
        return Read(stream, ReaderOptions.Default);
    }

    /// <inheritdoc />
    public DocumentModel Read(Stream stream, ReaderOptions options)
    {
        ArgumentNullException.ThrowIfNull(stream);
        ArgumentNullException.ThrowIfNull(options);

        options.Validate();

        // Check stream length before opening as zip
        if (stream.CanSeek && stream.Length > options.MaxUncompressedSize && options.MaxUncompressedSize > 0)
        {
            throw new DocumentTooLargeException(
                $"Document exceeds maximum uncompressed size limit of {options.MaxUncompressedSize} bytes (actual: {stream.Length} bytes).")
            {
                LimitType = DocumentTooLargeException.SizeLimitType.MaxUncompressedSize,
                MaxLimit = options.MaxUncompressedSize,
                ActualValue = stream.Length
            };
        }

        // Open as zip to check entry count and compression ratio before parsing
        if (stream.CanSeek)
        {
            stream.Position = 0;
            using var zipArchive = new ZipArchive(stream, ZipArchiveMode.Read, false);

            // Check entry count
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

            // Check compression ratio by sampling a few entries
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

        using var document = WordprocessingDocument.Open(stream, false);
        var body = document.MainDocumentPart?.Document.Body ?? throw new InvalidDataException("Document body missing.");

        // Check document part size
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
}
