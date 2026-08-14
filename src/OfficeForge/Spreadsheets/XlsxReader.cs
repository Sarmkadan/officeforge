using System.Globalization;
using System.IO;
using System.IO.Compression;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeForge.Models;

namespace OfficeForge.Spreadsheets;

public sealed class XlsxReader : IDocumentReader<WorkbookModel>
{
    /// <inheritdoc />
    public WorkbookModel Read(Stream stream)
    {
        return Read(stream, ReaderOptions.Default);
    }

    /// <inheritdoc />
    public WorkbookModel Read(Stream stream, ReaderOptions options)
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

        SpreadsheetDocument document;
        try
        {
            document = SpreadsheetDocument.Open(stream, false);
        }
        catch (OpenXmlPackageException ex)
        {
            throw new OfficeForgeFormatException("The document is not a valid Excel document.", ex);
        }

        using (document)
        {
            var workbookPart = document.WorkbookPart ?? throw new InvalidDataException("Workbook part missing.");

            // Check workbook part size
            try
            {
                using var partStream = workbookPart.GetStream();
                var partLength = partStream.Length;
                if (options.MaxPartSize > 0 && partLength > options.MaxPartSize)
                {
                    throw new DocumentTooLargeException(
                        $"Workbook part exceeds maximum size limit of {options.MaxPartSize} bytes (actual: {partLength} bytes).")
                    {
                        LimitType = DocumentTooLargeException.SizeLimitType.MaxPartSize,
                        MaxLimit = options.MaxPartSize,
                        ActualValue = partLength,
                        PartName = "Workbook"
                    };
                }
            }
            catch (IOException)
            {
                // Stream is not accessible, skip size check
            }

            var sharedStrings = workbookPart.SharedStringTablePart?.SharedStringTable?
                .Elements<SharedStringItem>().Select(i => i.InnerText).ToArray() ?? [];
            var model = new WorkbookModel();
            foreach (var sheet in workbookPart.Workbook.Descendants<Sheet>())
            {
                if (sheet.Id?.Value is not { } relId || sheet.Name?.Value is not { } name) continue;
                if (workbookPart.GetPartById(relId) is not WorksheetPart worksheetPart) continue;

                // Check worksheet part size
                try
                {
                    using var sheetStream = worksheetPart.GetStream();
                    var sheetLength = sheetStream.Length;
                    if (options.MaxPartSize > 0 && sheetLength > options.MaxPartSize)
                    {
                        throw new DocumentTooLargeException(
                            $"Worksheet '{name}' part exceeds maximum size limit of {options.MaxPartSize} bytes (actual: {sheetLength} bytes).")
                        {
                            LimitType = DocumentTooLargeException.SizeLimitType.MaxPartSize,
                            MaxLimit = options.MaxPartSize,
                            ActualValue = sheetLength,
                            PartName = $"Worksheet '{name}'"
                        };
                    }
                }
                catch (IOException)
                {
                    // Stream is not accessible, skip size check
                }

                var sheetModel = model.AddSheet(name);
                foreach (var cell in worksheetPart.Worksheet.Descendants<Cell>())
                {
                    if (cell.CellReference?.Value is not { } reference) continue;
                    if (!CellRef.TryParse(reference, out var cellRef)) continue;
                    var value = ConvertCell(cell, sharedStrings);
                    if (!value.IsEmpty) sheetModel[cellRef] = value;
                }
            }
            return model;
        }
    }

    /// <inheritdoc />
    public WorkbookModel Read(string path)
    {
        ArgumentNullException.ThrowIfNull(path);
        ArgumentException.ThrowIfNullOrEmpty(path);
        using var stream = File.OpenRead(path);
        return Read(stream);
    }

    /// <inheritdoc />
    public WorkbookModel Read(string path, ReaderOptions options)
    {
        ArgumentNullException.ThrowIfNull(path);
        ArgumentException.ThrowIfNullOrEmpty(path);
        ArgumentNullException.ThrowIfNull(options);

        options.Validate();

        using var stream = File.OpenRead(path);
        return Read(stream, options);
    }

    private static Models.CellValue ConvertCell(Cell cell, string[] sharedStrings)
    {
        if (cell.CellFormula?.Text is { Length: > 0 } formula)
            return Models.CellValue.FromFormula(formula);
        var raw = cell.CellValue?.InnerText;
        if (string.IsNullOrEmpty(raw) && cell.DataType?.Value != CellValues.InlineString)
            return Models.CellValue.Empty;
        var type = cell.DataType?.Value;
        if (type == CellValues.SharedString)
            return int.TryParse(raw, out var index) && index >= 0 && index < sharedStrings.Length
                ? Models.CellValue.FromText(sharedStrings[index])
                : Models.CellValue.Empty;
        if (type == CellValues.Boolean)
            return Models.CellValue.FromBoolean(raw == "1");
        if (type == CellValues.String || type == CellValues.InlineString)
            return Models.CellValue.FromText(type == CellValues.InlineString ? cell.InnerText : raw ?? string.Empty);
        if (type == CellValues.Date && DateTime.TryParse(raw, CultureInfo.InvariantCulture, DateTimeStyles.RoundtripKind, out var date))
            return Models.CellValue.FromDateTime(date);
        return double.TryParse(raw, NumberStyles.Float, CultureInfo.InvariantCulture, out var number)
            ? Models.CellValue.FromNumber(number)
            : Models.CellValue.FromText(raw ?? string.Empty);
    }
}
