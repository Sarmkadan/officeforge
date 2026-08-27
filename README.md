# officeforge

Cross-platform .NET library and CLI to create, edit and convert Word/Excel/PowerPoint documents, built on the OpenXML SDK.

Server-side Office document generation on Linux/CI without Office installed and without COM interop - the popular tools in this space are Windows-bound.

## Install

```bash
dotnet add package OfficeForge          # library
dotnet tool install -g OfficeForge.Cli  # CLI (command: officeforge)
```

## Quickstart

```csharp
using OfficeForge;
using OfficeForge.Export;
using OfficeForge.Models;

// Create a workbook, write cells, save
var workbook = new WorkbookModel();
var sheet = workbook.AddSheet("Report");
sheet["A1"] = CellValue.FromText("Revenue");
sheet["B1"] = CellValue.FromNumber(42500.75);
OfficeDocument.SaveWorkbook(workbook, "report.xlsx");

// Open it back and read a cell
var loaded = OfficeDocument.OpenWorkbook("report.xlsx");
Console.WriteLine(loaded.Sheets[0]["B1"]); // 42500.75

// Export any document to markdown
Console.WriteLine(OfficeDocument.Export("report.xlsx", ExportFormat.Markdown));
```

CLI:

```bash
officeforge write-cell report.xlsx B2 1999.5   # creates the file if missing
officeforge read-cell report.xlsx B2
officeforge convert report.xlsx --format markdown
```

## Features

| Area | Support |
| --- | --- |
| Excel (.xlsx) | read, write, cell-level editing, formulas |
| Word (.docx) | read, write (paragraphs, headings, tables) |
| PowerPoint (.pptx) | read (slides, titles, shape text) |
| Templates | `{{key}}` placeholder fill in .docx and .xlsx |
| Export | text, markdown, json from any supported format |

## XlsxRoundTripTests

`XlsxRoundTripTests` ([tests/OfficeForge.Tests/RoundTripTests.cs](tests/OfficeForge.Tests/RoundTripTests.cs)) is the round-trip test fixture for the Excel pipeline: it writes a workbook with the OfficeForge writers, loads it back with the readers and asserts nothing was lost along the way. It verifies that typed cell values (text, numbers, booleans, dates), multiple sheets and the default sheet of an empty workbook all survive a save/load cycle, that exporting a saved `.xlsx` path produces a markdown table, and that reading a missing cell yields an empty value. The fixture holds temporary files, so it implements `IDisposable` and should be disposed when done.

Usage:

```csharp
using OfficeForge.Tests;

// Run the xlsx round-trip checks against real files on disk
using var roundTrip = new XlsxRoundTripTests();

roundTrip.WriteRead_PreservesTypedCellValues();        // text/number/bool/date cells survive save + load
roundTrip.WriteRead_PreservesMultipleSheets();         // every sheet is written back and re-read intact
roundTrip.WriteRead_EmptyWorkbookGetsDefaultSheet();   // an empty workbook still round-trips with one sheet
roundTrip.Export_FromXlsxPath_ProducesMarkdownTable(); // Export(path, ExportFormat.Markdown) renders a table
roundTrip.MissingCell_ReadsAsEmpty();                  // absent cells read back as empty CellValue
```

## TemplateFillerTests

`TemplateFillerTests` (tests/OfficeForge.Tests/TemplateAndExportTests.cs) tests the template filling functionality for plain text, Word documents, and Excel workbooks, including placeholder replacement, parsing of key-value pairs, and export to various formats. It verifies that known placeholders are replaced, whitespace inside braces is tolerated, unknown placeholders are left intact, document and workbook filling preserves styles and untouched content, and that parsing handles malformed input correctly.

Usage:

```csharp
using OfficeForge.Templates;
using OfficeForge.Models;
using System.Collections.Generic;

// Example 1: Replace placeholders in plain text
var filler = new TemplateFiller(new Dictionary<string, string>
{
    { "name", "Ada" },
    { "company", "Acme" }
});
string filledText = filler.FillText("Hello {{name}} from {{company}}!");
// filledText is "Hello Ada from Acme!"

// Example 2: Fill a Word document (preserving styles and untouched paragraphs)
var document = new DocumentModel();
document.AddParagraph("Total: {{total}} EUR", new RunStyle { Bold = true });
document.AddParagraph("This paragraph has no placeholders and should remain unchanged.");
filler.Fill(document);
// The first paragraph now reads "Total: 99 EUR" and keeps its bold style.
// The second paragraph is unchanged.

// Example 3: Parse key-value pairs (used internally by TemplateFiller)
var pairs = TemplateFiller.ParsePairs(new[]
{
    "key=a=b",
    "other=x"
});
// pairs["key"] is "a=b", pairs["other"] is "x"
```

## Architecture

Readers map each format to a plain in-memory model (`WorkbookModel`, `DocumentModel`, `PresentationModel`); writers, exporters and the template filler operate on those models - the OpenXML SDK never leaks into the public API. See [docs/ARCHITECTURE.md](docs/ARCHITECTURE.md) for the module breakdown, design decisions and known limitations.

## Requirements

- .NET 10.0
- No Microsoft Office, no COM interop
- Linux, macOS, Windows

## License

MIT
