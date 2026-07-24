using System.IO;
using System.IO.Compression;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using OfficeForge.Models;
using A = DocumentFormat.OpenXml.Drawing;

namespace OfficeForge.Slides;

public sealed class PptxReader : IDocumentReader<PresentationModel>
{
    /// <inheritdoc />
    public PresentationModel Read(Stream stream)
    {
        return Read(stream, ReaderOptions.Default);
    }

    /// <inheritdoc />
    public PresentationModel Read(Stream stream, ReaderOptions options)
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

        using var document = PresentationDocument.Open(stream, false);
        var presentationPart = document.PresentationPart ?? throw new InvalidDataException("Presentation part missing.");

        // Check presentation part size
        try
        {
            using var partStream = presentationPart.GetStream();
            var partLength = partStream.Length;
            if (options.MaxPartSize > 0 && partLength > options.MaxPartSize)
            {
                throw new DocumentTooLargeException(
                    $"Presentation part exceeds maximum size limit of {options.MaxPartSize} bytes (actual: {partLength} bytes).")
                {
                    LimitType = DocumentTooLargeException.SizeLimitType.MaxPartSize,
                    MaxLimit = options.MaxPartSize,
                    ActualValue = partLength,
                    PartName = "Presentation"
                };
            }
        }
        catch (IOException)
        {
            // Stream is not accessible, skip size check
        }

        var model = new PresentationModel();
        var slideIds = presentationPart.Presentation.SlideIdList?.Elements<SlideId>() ?? [];
        foreach (var slideId in slideIds)
        {
            if (slideId.RelationshipId?.Value is not { } relId) continue;
            if (presentationPart.GetPartById(relId) is not SlidePart slidePart) continue;

            // Check slide part size
            if (options.MaxPartSize > 0)
            {
                var slideLength = slidePart.GetStream().Length;
                if (slideLength > options.MaxPartSize)
                {
                    throw new DocumentTooLargeException(
                        $"Slide part exceeds maximum size limit of {options.MaxPartSize} bytes (actual: {slideLength} bytes).")
                    {
                        LimitType = DocumentTooLargeException.SizeLimitType.MaxPartSize,
                        MaxLimit = options.MaxPartSize,
                        ActualValue = slideLength,
                        PartName = $"Slide {slideId.Id}"
                    };
                }
            }

            try
            {
                using var slideStream = slidePart.GetStream();
                var slideLength = slideStream.Length;
                if (options.MaxPartSize > 0 && slideLength > options.MaxPartSize)
                {
                    throw new DocumentTooLargeException(
                        $"Slide part exceeds maximum size limit of {options.MaxPartSize} bytes (actual: {slideLength} bytes).")
                    {
                        LimitType = DocumentTooLargeException.SizeLimitType.MaxPartSize,
                        MaxLimit = options.MaxPartSize,
                        ActualValue = slideLength,
                        PartName = $"Slide {slideId.Id}"
                    };
                }
            }
            catch (IOException)
            {
                // Stream is not accessible, skip size check
            }

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

    /// <inheritdoc />
    public PresentationModel Read(string path)
    {
        ArgumentNullException.ThrowIfNull(path);
        ArgumentException.ThrowIfNullOrEmpty(path);
        using var stream = File.OpenRead(path);
        return Read(stream);
    }

    /// <inheritdoc />
    public PresentationModel Read(string path, ReaderOptions options)
    {
        ArgumentNullException.ThrowIfNull(path);
        ArgumentException.ThrowIfNullOrEmpty(path);
        ArgumentNullException.ThrowIfNull(options);

        options.Validate();

        using var stream = File.OpenRead(path);
        return Read(stream, options);
    }
}
