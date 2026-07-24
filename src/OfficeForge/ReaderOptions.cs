using System;

namespace OfficeForge;

/// <summary>
/// Provides configuration options for document readers to guard against zip bombs and excessive resource consumption.
/// </summary>
public sealed class ReaderOptions
{
    /// <summary>
    /// Default options with conservative limits suitable for most server-side scenarios.
    /// </summary>
    public static readonly ReaderOptions Default = new ReaderOptions
    {
        MaxUncompressedSize = 50_000_000, // 50 MB
        MaxEntryCount = 10_000,
        MaxCompressionRatio = 100,
        MaxPartSize = 10_000_000 // 10 MB per part
    };

    /// <summary>
    /// Default options with more permissive limits for trusted environments.
    /// </summary>
    public static readonly ReaderOptions Permissive = new ReaderOptions
    {
        MaxUncompressedSize = 200_000_000, // 200 MB
        MaxEntryCount = 50_000,
        MaxCompressionRatio = 200,
        MaxPartSize = 50_000_000 // 50 MB per part
    };

    /// <summary>
    /// Gets or sets the maximum allowed uncompressed size of the entire document in bytes.
    /// A value of 0 or negative means no limit.
    /// Default: 50,000,000 (50 MB)
    /// </summary>
    public long MaxUncompressedSize { get; set; } 

    /// <summary>
    /// Gets or sets the maximum allowed number of entries in the zip archive.
    /// A value of 0 or negative means no limit.
    /// Default: 10,000
    /// </summary>
    public int MaxEntryCount { get; set; } 

    /// <summary>
    /// Gets or sets the maximum allowed compression ratio (uncompressed size / compressed size).
    /// A value of 0 or negative means no limit.
    /// Default: 100
    /// </summary>
    public int MaxCompressionRatio { get; set; } 

    /// <summary>
    /// Gets or sets the maximum allowed size of a single part within the document in bytes.
    /// A value of 0 or negative means no limit.
    /// Default: 10,000,000 (10 MB)
    /// </summary>
    public long MaxPartSize { get; set; } 

    /// <summary>
    /// Validates that the provided options are reasonable and won't cause immediate issues.
    /// </summary>
    /// <exception cref="ArgumentException">Thrown if options are invalid.</exception>
    public void Validate()
    {
        if (MaxUncompressedSize < 0)
        {
            throw new ArgumentException(
                $"{nameof(MaxUncompressedSize)} must be non-negative, but was {MaxUncompressedSize}.",
                nameof(MaxUncompressedSize));
        }

        if (MaxEntryCount < 0)
        {
            throw new ArgumentException(
                $"{nameof(MaxEntryCount)} must be non-negative, but was {MaxEntryCount}.",
                nameof(MaxEntryCount));
        }

        if (MaxCompressionRatio <= 0)
        {
            throw new ArgumentException(
                $"{nameof(MaxCompressionRatio)} must be positive, but was {MaxCompressionRatio}.",
                nameof(MaxCompressionRatio));
        }

        if (MaxPartSize < 0)
        {
            throw new ArgumentException(
                $"{nameof(MaxPartSize)} must be non-negative, but was {MaxPartSize}.",
                nameof(MaxPartSize));
        }
    }
}