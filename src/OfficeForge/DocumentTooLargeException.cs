using System;
using System.IO;
using System.Runtime.Serialization;

namespace OfficeForge;

/// <summary>
/// The exception that is thrown when a document or document part exceeds configured size limits.
/// This typically indicates a potential zip bomb or maliciously crafted document.
/// </summary>
[Serializable]
public sealed class DocumentTooLargeException : IOException
{
    /// <summary>
    /// Initializes a new instance of the <see cref="DocumentTooLargeException"/> class.
    /// </summary>
    public DocumentTooLargeException()
    {
    }

    /// <summary>
    /// Initializes a new instance of the <see cref="DocumentTooLargeException"/> class with a specified error message.
    /// </summary>
    /// <param name="message">The message that describes the error.</param>
    public DocumentTooLargeException(string? message) : base(message)
    {
    }

    /// <summary>
    /// Initializes a new instance of the <see cref="DocumentTooLargeException"/> class with a specified error message and a reference to the inner exception that is the cause of this exception.
    /// </summary>
    /// <param name="message">The error message that explains the reason for the exception.</param>
    /// <param name="innerException">The exception that is the cause of the current exception.</param>
    public DocumentTooLargeException(string? message, Exception? innerException) : base(message, innerException)
    {
    }

    /// <summary>
    /// Initializes a new instance of the <see cref="DocumentTooLargeException"/> class with serialized data.
    /// </summary>
    /// <param name="info">The <see cref="SerializationInfo"/> that holds the serialized object data about the exception being thrown.</param>
    /// <param name="context">The <see cref="StreamingContext"/> that contains contextual information about the source or destination.</param>
    // Obsolete serialization constructor - kept for binary serialization compatibility
    [Obsolete("This API supports obsolete formatter-based serialization and should not be used.")]
    private DocumentTooLargeException(SerializationInfo info, StreamingContext context)
        : base(info, context)
    {
        PartName = info.GetString(nameof(PartName));
        MaxLimit = info.GetInt64(nameof(MaxLimit));
        ActualValue = info.GetInt64(nameof(ActualValue));
        LimitType = (SizeLimitType)info.GetInt32(nameof(LimitType));
    }

    /// <inheritdoc />
    [Obsolete("This API supports obsolete formatter-based serialization and should not be used.")]
    public override void GetObjectData(SerializationInfo info, StreamingContext context)
    {
        base.GetObjectData(info, context);
        info.AddValue(nameof(PartName), PartName);
        info.AddValue(nameof(MaxLimit), MaxLimit);
        info.AddValue(nameof(ActualValue), ActualValue);
        info.AddValue(nameof(LimitType), (int)LimitType);
    }

    /// <summary>
    /// Gets the part of the document that caused the size violation, if available.
    /// </summary>
    public string? PartName { get; internal set; }

    /// <summary>
    /// Gets the maximum allowed value for the violated limit.
    /// </summary>
    public long MaxLimit { get; internal set; }

    /// <summary>
    /// Gets the actual value that exceeded the limit.
    /// </summary>
    public long ActualValue { get; internal set; }

    /// <summary>
    /// Gets the type of size limit that was violated.
    /// </summary>
    public SizeLimitType LimitType { get; internal set; }

    /// <summary>
    /// The type of size limit being enforced.
    /// </summary>
    public enum SizeLimitType
    {
        /// <summary>Maximum uncompressed size of the entire document.</summary>
        MaxUncompressedSize,
        /// <summary>Maximum number of entries in the zip archive.</summary>
        MaxEntryCount,
        /// <summary>Maximum compression ratio (uncompressed/compressed size).</summary>
        MaxCompressionRatio,
        /// <summary>Maximum size of a single part within the document.</summary>
        MaxPartSize
    }
}