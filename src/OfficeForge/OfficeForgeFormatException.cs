using System;

namespace OfficeForge;

/// <summary>
/// Thrown when an Office document cannot be parsed due to format errors.
/// </summary>
public sealed class OfficeForgeFormatException : Exception
{
    /// <summary>
    /// Initializes a new instance of the <see cref="OfficeForgeFormatException"/> class.
    /// </summary>
    /// <param name="message">The error message.</param>
    public OfficeForgeFormatException(string message) : base(message) { }

    /// <summary>
    /// Initializes a new instance of the <see cref="OfficeForgeFormatException"/> class.
    /// </summary>
    /// <param name="message">The error message.</param>
    /// <param name="innerException">The exception that caused this one.</param>
    public OfficeForgeFormatException(string message, Exception innerException) : base(message, innerException) { }
}
