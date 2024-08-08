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
| PowerPoint (.pptx) | read (slides, text, notes) |
| Templates | `{{key}}` placeholder fill in .docx and .xlsx |
| Export | text, markdown, json from any supported format |

## Requirements

- .NET 10.0
- No Microsoft Office, no COM interop
- Linux, macOS, Windows

## License

MIT
