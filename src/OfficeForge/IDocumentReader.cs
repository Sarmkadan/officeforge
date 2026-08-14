using System.IO;

namespace OfficeForge;

public interface IDocumentReader<TModel>
{
    /// <summary>
    /// Reads a document from a stream.
    /// </summary>
    /// <param name="stream">The stream containing the document data.</param>
    /// <returns>The parsed document model.</returns>
    /// <exception cref="ArgumentNullException"><paramref name="stream"/> is <see langword="null"/></exception>
    /// <exception cref="DocumentTooLargeException">The document exceeds configured size limits</exception>
    /// <exception cref="OfficeForgeFormatException">The document is not a valid OpenXML package</exception>
    TModel Read(Stream stream);

    /// <summary>
    /// Reads a document from a stream with custom reader options.
    /// </summary>
    /// <param name="stream">The stream containing the document data.</param>
    /// <param name="options">Reader configuration options.</param>
    /// <returns>The parsed document model.</returns>
    /// <exception cref="ArgumentNullException"><paramref name="stream"/> or <paramref name="options"/> is <see langword="null"/></exception>
    /// <exception cref="DocumentTooLargeException">The document exceeds configured size limits</exception>
    /// <exception cref="OfficeForgeFormatException">The document is not a valid OpenXML package</exception>
    TModel Read(Stream stream, ReaderOptions options);

    /// <summary>
    /// Reads a document from a file path.
    /// </summary>
    /// <param name="path">The file path to read from.</param>
    /// <returns>The parsed document model.</returns>
    /// <exception cref="ArgumentNullException"><paramref name="path"/> is <see langword="null"/></exception>
    /// <exception cref="ArgumentException"><paramref name="path"/> is empty or whitespace</exception>
    /// <exception cref="FileNotFoundException">The file does not exist</exception>
    /// <exception cref="IOException">The file could not be read</exception>
    /// <exception cref="DocumentTooLargeException">The document exceeds configured size limits</exception>
    /// <exception cref="OfficeForgeFormatException">The document is not a valid OpenXML package</exception>
    TModel Read(string path);

    /// <summary>
    /// Reads a document from a file path with custom reader options.
    /// </summary>
    /// <param name="path">The file path to read from.</param>
    /// <param name="options">Reader configuration options.</param>
    /// <returns>The parsed document model.</returns>
    /// <exception cref="ArgumentNullException"><paramref name="path"/> or <paramref name="options"/> is <see langword="null"/></exception>
    /// <exception cref="ArgumentException"><paramref name="path"/> is empty or whitespace</exception>
    /// <exception cref="FileNotFoundException">The file does not exist</exception>
    /// <exception cref="IOException">The file could not be read</exception>
    /// <exception cref="DocumentTooLargeException">The document exceeds configured size limits</exception>
    /// <exception cref="OfficeForgeFormatException">The document is not a valid OpenXML package</exception>
    TModel Read(string path, ReaderOptions options);
}

public interface IDocumentWriter<TModel>
{
    /// <summary>
    /// Writes a document to a stream.
    /// </summary>
    /// <param name="model">The document model to write.</param>
    /// <param name="stream">The stream to write to.</param>
    /// <exception cref="ArgumentNullException"><paramref name="model"/> or <paramref name="stream"/> is <see langword="null"/></exception>
    void Write(TModel model, Stream stream);

    /// <summary>
    /// Writes a document to a file path.
    /// </summary>
    /// <param name="model">The document model to write.</param>
    /// <param name="path">The file path to write to.</param>
    /// <exception cref="ArgumentNullException"><paramref name="model"/> is <see langword="null"/></exception>
    /// <exception cref="ArgumentException"><paramref name="path"/> is null, empty, or whitespace</exception>
    void Write(TModel model, string path);
}
