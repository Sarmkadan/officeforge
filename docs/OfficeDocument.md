# OfficeDocument

`OfficeDocument` is the central static facade of the **officeforge** library. It provides a single entry point for detecting file formats, opening Microsoft Office files into strongly typed in-memory models, saving modified models back to disk, and exporting document content to various text-based representations. All operations are synchronous and stateless.

## API

### DetectKind

```csharp
public static DocumentKind DetectKind(string filePath)
```

Determines the document kind of the file at the given path by inspecting its structure or signature, without fully parsing its content.

- **Parameters**: `filePath` — the absolute or relative path to the file.
- **Returns**: a `DocumentKind` enumeration value (`Workbook`, `Document`, `Presentation`, or `Unknown`).
- **Throws**:
  - `FileNotFoundException` if the file does not exist.
  - `InvalidOperationException` if the file is accessible but its format cannot be read at all (e.g., corrupt header).

### OpenWorkbook

```csharp
public static WorkbookModel OpenWorkbook(string filePath)
```

Opens an Excel workbook (`.xlsx`, `.xlsm`, or similar) and returns its complete in-memory representation.

- **Parameters**: `filePath` — path to the workbook file.
- **Returns**: a populated `WorkbookModel` containing sheets, rows, cells, and metadata.
- **Throws**:
  - `FileNotFoundException` if the file does not exist.
  - `InvalidDataException` if the file is not a valid workbook format.
  - `NotSupportedException` if the workbook uses features not supported by the reader (e.g., legacy binary `.xls`).

### OpenDocument

```csharp
public static DocumentModel OpenDocument(string filePath)
```

Opens a Word processing document (`.docx`) and returns its in-memory model.

- **Parameters**: `filePath` — path to the document file.
- **Returns**: a `DocumentModel` containing paragraphs, runs, styles, and section properties.
- **Throws**:
  - `FileNotFoundException` if the file does not exist.
  - `InvalidDataException` if the file is not a valid DOCX package.

### OpenPresentation

```csharp
public static PresentationModel OpenPresentation(string filePath)
```

Opens a PowerPoint presentation (`.pptx`) and returns its in-memory model.

- **Parameters**: `filePath` — path to the presentation file.
- **Returns**: a `PresentationModel` containing slides, shapes, and layout information.
- **Throws**:
  - `FileNotFoundException` if the file does not exist.
  - `InvalidDataException` if the file is not a valid PPTX package.

### SaveWorkbook

```csharp
public static void SaveWorkbook(WorkbookModel model, string filePath)
```

Writes a `WorkbookModel` back to disk as an Office Open XML workbook file.

- **Parameters**:
  - `model` — the workbook model to persist; must not be null.
  - `filePath` — the destination file path; any existing file is overwritten.
- **Throws**:
  - `ArgumentNullException` if `model` is null.
  - `DirectoryNotFoundException` if the target directory does not exist.
  - `IOException` if the file is locked or the path is invalid.

### SaveDocument

```csharp
public static void SaveDocument(DocumentModel model, string filePath)
```

Writes a `DocumentModel` back to disk as a DOCX file.

- **Parameters**:
  - `model` — the document model to persist; must not be null.
  - `filePath` — the destination file path; any existing file is overwritten.
- **Throws**:
  - `ArgumentNullException` if `model` is null.
  - `DirectoryNotFoundException` if the target directory does not exist.
  - `IOException` if the file is locked or the path is invalid.

### Export

```csharp
public static string Export(object model, string format)
```

Converts any supported document model into a string representation according to the specified format.

- **Parameters**:
  - `model` — a `WorkbookModel`, `DocumentModel`, or `PresentationModel` instance; must not be null.
  - `format` — a case-insensitive format identifier. Supported values: `"text"`, `"markdown"`, `"json"`.
- **Returns**: the full exported content as a string.
- **Throws**:
  - `ArgumentNullException` if `model` is null.
  - `ArgumentException` if `format` is null, empty, or not one of the supported identifiers.
  - `NotSupportedException` if the model type does not have an exporter for the requested format.

## Usage

### Example 1: Detect, open, modify, and save a workbook

```csharp
string path = @"C:\Reports\quarterly.xlsx";

DocumentKind kind = OfficeDocument.DetectKind(path);
if (kind != DocumentKind.Workbook)
{
    Console.WriteLine("Unsupported file type.");
    return;
}

WorkbookModel workbook = OfficeDocument.OpenWorkbook(path);

// Modify a cell value
workbook.Sheets[0].Rows[0].Cells[0].Value = "Updated Header";

OfficeDocument.SaveWorkbook(workbook, path);
```

### Example 2: Open a document and export it to Markdown

```csharp
string docPath = @"C:\Docs\proposal.docx";

DocumentModel document = OfficeDocument.OpenDocument(docPath);

string markdown = OfficeDocument.Export(document, "markdown");
File.WriteAllText(@"C:\Docs\proposal.md", markdown);
```

## Notes

- **Format detection**: `DetectKind` relies on file signatures and package structure, not file extensions. A `.docx` renamed to `.zip` will still be identified as `Document`. Corrupt or truncated files may yield `Unknown`.
- **Export format availability**: The `"text"` and `"markdown"` exporters are available for all three model types. The `"json"` exporter serializes the full model tree and is intended for debugging or interchange, not for round-tripping back to Office formats.
- **Overwrite behavior**: `SaveWorkbook` and `SaveDocument` overwrite the target file silently. Callers must ensure the path is writable and that overwriting is intended.
- **Thread safety**: All methods on `OfficeDocument` are static and operate on caller-supplied paths and models. No shared mutable state is held internally. Multiple threads may safely call these methods concurrently, provided they do not pass the same model instance to mutating operations without external synchronization. The model objects themselves are not thread-safe; concurrent modification of a single `WorkbookModel` or `DocumentModel` instance across threads will produce undefined behavior.
- **Legacy formats**: Binary `.xls`, `.doc`, and `.ppt` formats are not supported and will cause `OpenWorkbook`, `OpenDocument`, or `OpenPresentation` to throw `NotSupportedException` or `InvalidDataException`. Use `DetectKind` first to guard against unsupported input.
