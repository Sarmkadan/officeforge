using OfficeForge.Export;
using OfficeForge.Models;
using OfficeForge.Slides;
using OfficeForge.Spreadsheets;
using OfficeForge.Words;

namespace OfficeForge;

public enum DocumentKind
{
    Workbook,
    Document,
    Presentation
}

public static class OfficeDocument
{
    /// <summary>
    /// Detects the document kind based on the file extension.
    /// </summary>
    /// <param name="path">The file path.</param>
    /// <returns>The detected <see cref="DocumentKind"/>.</returns>
    /// <exception cref="ArgumentException">Thrown when <paramref name="path"/> is null, empty or whitespace.</exception>
    /// <exception cref="NotSupportedException">Thrown when the extension is not recognised.</exception>
    public static DocumentKind DetectKind(string path)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(path);
        return Path.GetExtension(path).ToLowerInvariant() switch
        {
            ".xlsx" or ".xlsm" => DocumentKind.Workbook,
            ".docx" => DocumentKind.Document,
            ".pptx" => DocumentKind.Presentation,
            var ext => throw new NotSupportedException($"Unsupported file extension '{ext}'.")
        };
    }

    /// <summary>
    /// Opens a workbook from the specified path.
    /// </summary>
    /// <param name="path">The workbook file path.</param>
    /// <returns>The loaded <see cref="WorkbookModel"/>.</returns>
    public static WorkbookModel OpenWorkbook(string path) => new XlsxReader().Read(path);

    /// <summary>
    /// Opens a document from the specified path.
    /// </summary>
    /// <param name="path">The document file path.</param>
    /// <returns>The loaded <see cref="DocumentModel"/>.</returns>
    public static DocumentModel OpenDocument(string path) => new DocxReader().Read(path);

    /// <summary>
    /// Opens a presentation from the specified path.
    /// </summary>
    /// <param name="path">The presentation file path.</param>
    /// <returns>The loaded <see cref="PresentationModel"/>.</returns>
    public static PresentationModel OpenPresentation(string path) => new PptxReader().Read(path);

    /// <summary>
    /// Opens a workbook from the specified path using custom <paramref name="options"/>.
    /// </summary>
    /// <param name="path">The workbook file path.</param>
    /// <param name="options">Reader options.</param>
    /// <returns>The loaded <see cref="WorkbookModel"/>.</returns>
    public static WorkbookModel OpenWorkbook(string path, ReaderOptions options) => new XlsxReader().Read(path, options);

    /// <summary>
    /// Opens a document from the specified path using custom <paramref name="options"/>.
    /// </summary>
    /// <param name="path">The document file path.</param>
    /// <param name="options">Reader options.</param>
    /// <returns>The loaded <see cref="DocumentModel"/>.</returns>
    public static DocumentModel OpenDocument(string path, ReaderOptions options) => new DocxReader().Read(path, options);

    /// <summary>
    /// Opens a presentation from the specified path using custom <paramref name="options"/>.
    /// </summary>
    /// <param name="path">The presentation file path.</param>
    /// <param name="options">Reader options.</param>
    /// <returns>The loaded <see cref="PresentationModel"/>.</returns>
    public static PresentationModel OpenPresentation(string path, ReaderOptions options) => new PptxReader().Read(path, options);

    /// <summary>
    /// Saves a workbook to the specified path.
    /// </summary>
    /// <param name="workbook">The workbook to save.</param>
    /// <param name="path">The destination path.</param>
    public static void SaveWorkbook(WorkbookModel workbook, string path)
    {
        ArgumentNullException.ThrowIfNull(workbook);
        new XlsxWriter().Write(workbook, path);
    }

    /// <summary>
    /// Saves a document to the specified path.
    /// </summary>
    /// <param name="document">The document to save.</param>
    /// <param name="path">The destination path.</param>
    public static void SaveDocument(DocumentModel document, string path)
    {
        ArgumentNullException.ThrowIfNull(document);
        new DocxWriter().Write(document, path);
    }

    /// <summary>
    /// Exports a file at <paramref name="path"/> to the requested <paramref name="format"/>.
    /// </summary>
    /// <param name="path">The source file path.</param>
    /// <param name="format">The desired export format.</param>
    /// <returns>The exported string.</returns>
    /// <exception cref="ArgumentException">Thrown when <paramref name="path"/> is null, empty or whitespace.</exception>
    /// <exception cref="NotSupportedException">Thrown when the file extension is unsupported.</exception>
    public static string Export(string path, ExportFormat format) => DetectKind(path) switch
    {
        DocumentKind.Workbook => WorkbookExporter.Export(path, format),
        DocumentKind.Document => DocumentExporter.Export(path, format),
        _ => PresentationExporter.Export(path, format)
    };
}
