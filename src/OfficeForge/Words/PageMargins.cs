namespace OfficeForge.Words;

/// <summary>
/// Page margins for a Word document, expressed in points (1/72 inch).
/// </summary>
public sealed class PageMargins
{
    /// <summary>
    /// Gets or sets the top margin in points. Default: 72 (1 inch).
    /// </summary>
    public double Top { get; set; } = 72;

    /// <summary>
    /// Gets or sets the bottom margin in points. Default: 72 (1 inch).
    /// </summary>
    public double Bottom { get; set; } = 72;

    /// <summary>
    /// Gets or sets the left margin in points. Default: 72 (1 inch).
    /// </summary>
    public double Left { get; set; } = 72;

    /// <summary>
    /// Gets or sets the right margin in points. Default: 72 (1 inch).
    /// </summary>
    public double Right { get; set; } = 72;
}
