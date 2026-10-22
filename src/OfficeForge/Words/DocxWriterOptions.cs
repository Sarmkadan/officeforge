using System;
using System.IO.Compression;

namespace OfficeForge.Words;

/// <summary>
/// Provides configuration options for <see cref="DocxWriter"/>: document-wide font and page layout defaults
/// and the ZIP compression level used for the package.
/// </summary>
public sealed class DocxWriterOptions
{
    /// <summary>
    /// Largest page margin accepted, in points (22 inches), the largest page dimension Word supports.
    /// </summary>
    public const double MaxMarginPoints = 1584;

    /// <summary>
    /// Default options: Calibri 11 pt, 1-inch margins, optimal compression.
    /// </summary>
    public static readonly DocxWriterOptions Default = new DocxWriterOptions();

    /// <summary>
    /// Gets or sets the default font family applied to document text that does not specify its own font.
    /// Must not be null, empty or whitespace. Default: <c>Calibri</c>.
    /// </summary>
    public string DefaultFontName { get; set; } = "Calibri";

    /// <summary>
    /// Gets or sets the default font size in points. Must be within 1 to 400 inclusive. Default: 11.
    /// </summary>
    public double DefaultFontSize { get; set; } = 11;

    /// <summary>
    /// Gets or sets the page margins in points. Must not be <c>null</c>. Default: 1 inch on every side.
    /// </summary>
    public PageMargins PageMargins { get; set; } = new PageMargins();

    /// <summary>
    /// Gets or sets the compression level used for the DOCX package. Default: <see cref="CompressionLevel.Optimal"/>.
    /// </summary>
    public CompressionLevel CompressionLevel { get; set; } = CompressionLevel.Optimal;

    /// <summary>
    /// Validates the options.
    /// </summary>
    /// <exception cref="ArgumentException">
    /// Thrown when a property holds an invalid value. The <see cref="ArgumentException.ParamName"/> is the name of the offending property.
    /// </exception>
    public void Validate()
    {
        if (string.IsNullOrWhiteSpace(DefaultFontName))
        {
            throw new ArgumentException(
                $"{nameof(DefaultFontName)} must be a non-empty font name.",
                nameof(DefaultFontName));
        }

        if (!(DefaultFontSize >= 1 && DefaultFontSize <= 400))
        {
            throw new ArgumentException(
                $"{nameof(DefaultFontSize)} must be between 1 and 400 points, but was {DefaultFontSize}.",
                nameof(DefaultFontSize));
        }

        if (PageMargins is null)
        {
            throw new ArgumentException($"{nameof(PageMargins)} must not be null.", nameof(PageMargins));
        }

        ValidateMargin(nameof(PageMargins.Top), PageMargins.Top);
        ValidateMargin(nameof(PageMargins.Bottom), PageMargins.Bottom);
        ValidateMargin(nameof(PageMargins.Left), PageMargins.Left);
        ValidateMargin(nameof(PageMargins.Right), PageMargins.Right);

        if (!Enum.IsDefined(CompressionLevel))
        {
            throw new ArgumentException(
                $"{nameof(CompressionLevel)} must be a defined {nameof(System.IO.Compression.CompressionLevel)} value, but was {(int)CompressionLevel}.",
                nameof(CompressionLevel));
        }
    }

    private static void ValidateMargin(string name, double value)
    {
        if (!(value >= 0 && value <= MaxMarginPoints))
        {
            throw new ArgumentException(
                $"PageMargins.{name} must be between 0 and {MaxMarginPoints} points, but was {value}.",
                $"{nameof(PageMargins)}.{name}");
        }
    }
}
