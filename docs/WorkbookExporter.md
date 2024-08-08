# WorkbookExporter

The `WorkbookExporter` class provides a set of static utility methods for serializing `officeforge` workbook objects into standardized text-based formats. It serves as a centralized interface for converting spreadsheet structures into human-readable representations or machine-interchangeable data formats.

## API

### Export
Serializes a workbook or a workbook file into a string representation.
- **Return Value**: A `string` containing the serialized representation of the workbook.
- **Throws**: May throw `ArgumentNullException` if the input workbook or path is null, or `IOException` if the file cannot be accessed.

### ToPlainText
Converts the content of a workbook into a plain text format, preserving the structure of cells and rows as text.
- **Return Value**: A `string` containing the plain text representation.

### ToMarkdown
Converts the workbook data into a Markdown table format suitable for documentation.
- **Return Value**: A `string` containing the Markdown table.

### ToJson
Serializes the workbook's content into a standard JSON string for data interchange.
- **Return Value**: A `string` containing the JSON representation of the workbook data.

## Usage

```csharp
using OfficeForge;

// Load a workbook
var workbook = Workbook.Load("data.xlsx");

// Example 1: Exporting to Markdown for documentation
string markdown = WorkbookExporter.ToMarkdown(workbook);
System.IO.File.WriteAllText("data.md", markdown);

// Example 2: Serializing to JSON for web API consumption
string json = WorkbookExporter.ToJson(workbook);
System.Console.WriteLine(json);
```

## Notes

- **Thread Safety**: All methods within `WorkbookExporter` are static and stateless, making them thread-safe for concurrent operations on separate `Workbook` instances.
- **Encoding**: Output strings are generated using default UTF-8 encoding.
- **Null Handling**: Methods will throw `ArgumentNullException` if the provided `Workbook` instance is null.
- **Data Size**: Extremely large workbooks may consume significant memory during serialization; callers should ensure adequate heap space is available when processing large spreadsheet files.
