using System;

namespace OfficeForge.Slides;

/// <summary>
/// Provides configuration options for <see cref="PptxWriter"/>: slide dimensions, the default theme name,
/// the slide count limit and whether speaker notes are included.
/// </summary>
public sealed class PptxWriterOptions
{
    /// <summary>
    /// Smallest slide dimension accepted, in EMU (1 inch), the lower bound of the OOXML <c>ST_SlideSizeCoordinate</c> type.
    /// </summary>
    public const int MinSlideDimensionEmu = 914400;

    /// <summary>
    /// Largest slide dimension accepted, in EMU (56 inches), the upper bound of the OOXML <c>ST_SlideSizeCoordinate</c> type.
    /// </summary>
    public const int MaxSlideDimensionEmu = 51206400;

    /// <summary>
    /// Default options: 4:3 slides (10 x 7.5 inches), the size a presentation without an explicit slide size implies,
    /// the default theme name, up to 1000 slides and no speaker notes.
    /// </summary>
    public static readonly PptxWriterOptions Default = new PptxWriterOptions();

    /// <summary>
    /// Gets or sets the slide width in EMU. Must be within 914400 to 51206400 inclusive. Default: 9144000 (10 inches).
    /// </summary>
    public int SlideWidth { get; set; } = 9144000;

    /// <summary>
    /// Gets or sets the slide height in EMU. Must be within 914400 to 51206400 inclusive. Default: 6858000 (7.5 inches).
    /// </summary>
    public int SlideHeight { get; set; } = 6858000;

    /// <summary>
    /// Gets or sets the name of the default theme. Must not be <c>null</c>, empty or whitespace. Default: <c>Office Theme</c>.
    /// </summary>
    public string DefaultTheme { get; set; } = "Office Theme";

    /// <summary>
    /// Gets or sets the maximum number of slides a presentation may contain. Must be positive. Default: 1000.
    /// </summary>
    public int MaxSlides { get; set; } = 1000;

    /// <summary>
    /// Gets or sets a value indicating whether speaker notes are included in the output. Default: <c>false</c>.
    /// </summary>
    public bool IncludeNotes { get; set; }

    /// <summary>
    /// Validates the options.
    /// </summary>
    /// <exception cref="ArgumentException">
    /// Thrown when a property holds an invalid value. The <see cref="ArgumentException.ParamName"/> is the name of the offending property.
    /// </exception>
    public void Validate()
    {
        ValidateDimension(nameof(SlideWidth), SlideWidth);
        ValidateDimension(nameof(SlideHeight), SlideHeight);

        if (string.IsNullOrWhiteSpace(DefaultTheme))
        {
            throw new ArgumentException(
                $"{nameof(DefaultTheme)} must be a non-empty theme name.",
                nameof(DefaultTheme));
        }

        if (MaxSlides <= 0)
        {
            throw new ArgumentException(
                $"{nameof(MaxSlides)} must be positive, but was {MaxSlides}.",
                nameof(MaxSlides));
        }
    }

    private static void ValidateDimension(string name, int value)
    {
        if (value < MinSlideDimensionEmu || value > MaxSlideDimensionEmu)
        {
            throw new ArgumentException(
                $"{name} must be between {MinSlideDimensionEmu} and {MaxSlideDimensionEmu} EMU, but was {value}.",
                name);
        }
    }
}
