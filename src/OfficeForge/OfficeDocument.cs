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
    public static DocumentKind DetectKind(string path) => Path.GetExtension(path).ToLowerInvariant() switch
    {
        ".xlsx" or ".xlsm" => DocumentKind.Workbook,
        ".docx" => DocumentKind.Document,
        ".pptx" => DocumentKind.Presentation,
        var ext => throw new NotSupportedException($"Unsupported file extension '{ext}'.")
    };

    public static WorkbookModel OpenWorkbook(string path) => new XlsxReader().Read(path);
    public static DocumentModel OpenDocument(string path) => new DocxReader().Read(path);
    public static PresentationModel OpenPresentation(string path) => new PptxReader().Read(path);
    public static void SaveWorkbook(WorkbookModel workbook, string path) => new XlsxWriter().Write(workbook, path);
    public static void SaveDocument(DocumentModel document, string path) => new DocxWriter().Write(document, path);

    public static string Export(string path, ExportFormat format) => DetectKind(path) switch
    {
        DocumentKind.Workbook => WorkbookExporter.Export(OpenWorkbook(path), format),
        DocumentKind.Document => DocumentExporter.Export(OpenDocument(path), format),
        _ => PresentationExporter.Export(OpenPresentation(path), format)
    };
}
