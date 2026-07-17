# Architecture

## Overview

OfficeForge is a small library (plus a thin CLI) for reading, writing and
converting Office Open XML documents (.xlsx/.xlsm, .docx, .pptx) without
Microsoft Office. It sits on top of the `DocumentFormat.OpenXml` SDK and
`System.IO.Packaging`, and deliberately does **not** expose the OpenXML object
model to callers. Instead, every format is mapped to a simple in-memory model,
and all reading/writing goes through those models.

The core idea:

```
 file/stream --(Reader)--> plain C# model --(Writer)--> file/stream
                                 |
                                 +--(Exporter)--> text / markdown / json
                                 +--(TemplateFiller)--> mutated model
```

Readers, writers, exporters and the template filler are all stateless; the
only mutable state lives in the models themselves.

## Projects

| Project | Purpose |
| --- | --- |
| `src/OfficeForge` | The library: models, readers, writers, exporters, templates |
| `src/OfficeForge.Cli` | `officeforge` command-line tool (System.CommandLine) built on the library |
| `tests/OfficeForge.Tests` | xUnit tests: round-trip (write then read back) and template/export behavior |

## Module breakdown (src/OfficeForge)

### Models (`Models/`)

Format-agnostic, dependency-free value objects. Nothing in `Models/`
references the OpenXML SDK.

- `WorkbookModel` / `SheetModel` - a workbook is a list of named sheets; a
  sheet is a **sparse** `Dictionary<CellRef, CellValue>`. Indexers accept
  either a `CellRef` or an A1 string (`sheet["B2"]`). Reading a missing cell
  returns `CellValue.Empty`, never throws. `OrderedCells()` yields cells in
  row-major order and is what writers/exporters iterate.
- `CellRef` - readonly record struct holding 1-based `(Row, Column)`, with
  A1 parsing/formatting (`Parse`, `TryParse`, `ToA1`, `ColumnName`).
- `CellValue` - readonly record struct tagged by `CellValueKind`
  (Empty/Text/Number/Boolean/DateTime/Formula). Constructed only via the
  `From*` factory methods; `ToString()` renders invariant-culture text
  (formulas render as `=FORMULA`).
- `DocumentModel` / `ParagraphModel` / `RunModel` / `RunStyle` - a document is
  a flat list of paragraphs; a paragraph is a list of styled runs plus a
  `ParagraphKind` (Body/Heading/ListItem) and `HeadingLevel`.
- `PresentationModel` / `SlideModel` / `ShapeTextModel` - text-only view of a
  deck: slides with an optional title and text lines grouped per shape.

### Readers and writers (`Spreadsheets/`, `Words/`, `Slides/`)

Each implements `IDocumentReader<TModel>` / `IDocumentWriter<TModel>`
(`IDocumentReader.cs`), with both `Stream` and `string path` overloads (the
path overloads just open a `FileStream` and delegate).

- `XlsxReader` - resolves shared strings up front, then converts each cell by
  its `DataType`: shared string, boolean, inline string, ISO date, otherwise
  invariant-culture number with a fallback to text. Cells whose reference
  fails `CellRef.TryParse` or whose value is empty are skipped, keeping the
  model sparse. A cell with a formula is captured as `CellValueKind.Formula`
  (the cached result is discarded - see limitations).
- `XlsxWriter` - emits one `WorksheetPart` per sheet (an empty workbook gets a
  default `Sheet1`, since a sheetless xlsx is invalid), grouping cells by row.
  Text is written as **inline strings** rather than a shared-string table -
  simpler, valid, at the cost of file size on highly repetitive data.
- `DocxReader` - flattens the body into paragraphs, detecting headings via the
  `Heading{N}` paragraph style id. Run formatting honors OnOff semantics:
  `<w:b/>` is bold but `<w:b w:val="0"/>` is not, and `<w:u w:val="none"/>` is
  not underlined.
- `DocxWriter` - mirrors the reader: `Heading{N}` style ids for heading
  paragraphs (levels 1-9), Bold/Italic/Underline/RunFonts run properties, and
  `xml:space="preserve"` on every text node so leading/trailing whitespace
  survives.
- `PptxReader` - read-only. Walks `SlideIdList` in presentation order, and per
  slide extracts non-empty text lines per shape. The first Title/CenteredTitle
  placeholder becomes `SlideModel.Title`; everything else becomes a
  `ShapeTextModel`. There is no `PptxWriter` (see limitations).

### Facade (`OfficeDocument.cs`)

Static convenience API used by the CLI and most callers:
`DetectKind(path)` (by file extension), `OpenWorkbook`/`OpenDocument`/
`OpenPresentation`, `SaveWorkbook`/`SaveDocument`, and `Export(path, format)`
which detects the kind, opens the file and dispatches to the right exporter.
The facade instantiates readers/writers per call - they are cheap and
stateless, so no caching is needed.

### Export (`Export/Exporters.cs`)

Three static exporters (`WorkbookExporter`, `DocumentExporter`,
`PresentationExporter`), each supporting `ExportFormat.Text | Markdown | Json`:

- Workbook: text = tab-separated rows per sheet; markdown = one `##` section
  and pipe table per sheet (first row is the header, `|` escaped in cells);
  json = `{ sheetName: { "A1": "value", ... } }`.
- Document: markdown maps headings to `#`-prefixes (clamped to 6), list items
  to `- `, and bold/italic runs to `**`/`*` markers.
- Presentation: markdown = `## <title or "Slide N">` plus a bullet per line.

All JSON uses a single shared `JsonSerializerOptions` (`JsonOptions.Indented`),
and cell values are exported as their string rendering, not typed JSON values.

### Templates (`Templates/TemplateFiller.cs`)

Replaces `{{key}}` placeholders (whitespace-tolerant, key charset `[\w.-]`,
compile-time `GeneratedRegex`). Unknown keys are left untouched so partial
fills are safe. Two targets:

- `Fill(DocumentModel)`: operates on the concatenated paragraph text, so a
  placeholder split across runs (common after editing in Word) is still found.
  Trade-off: a paragraph that changes is collapsed into a single run carrying
  the *first* run's style.
- `Fill(WorkbookModel)`: per-cell, text cells only.

`ParsePairs` turns CLI-style `key=value` strings into the value dictionary.

### CLI (`src/OfficeForge.Cli/Program.cs`)

Top-level program using System.CommandLine. Commands: `read-cell`,
`write-cell` (creates the workbook/sheet if missing), `convert` (any format,
stdout or `--output`), `extract-text`, `fill-template` (`--set key=value`,
docx/xlsx only). All logic delegates to the `OfficeDocument` facade; the CLI
adds only argument parsing, `=`-prefix formula / bool / invariant-number value
coercion (`ParseValue`), and sheet resolution.

## Key design decisions

- **Own models instead of exposing OpenXML types.** The OpenXML SDK is a
  faithful but verbose XML mapping. Wrapping it lets callers work with
  `sheet["A1"] = CellValue.FromNumber(1)` and keeps the SDK an internal
  dependency. The cost is fidelity: anything the model doesn't represent
  (styles, merged cells, images, charts) is dropped on a read/write round
  trip. This is a lossy converter by design, not an editor that preserves
  unknown content.
- **Sparse cell storage.** Sheets are dictionaries, not 2D arrays, so huge
  row/column indices cost nothing. `RowCount`/`ColumnCount` are O(n) maxima
  over the keys, computed on demand.
- **Structs for `CellRef`/`CellValue`.** Both are small readonly record
  structs: value equality for dictionary keys, no allocation per cell value,
  and `CellValue` is immutable so it can be shared freely.
- **Static exporters/facade, instance readers/writers.** Readers/writers
  implement the generic interfaces so they can be mocked or swapped;
  exporters are pure string functions with no state to abstract.
- **Invariant culture everywhere** numbers/dates are parsed or formatted -
  xlsx content is culture-independent by spec.

## Data flow example (`officeforge convert report.xlsx --format markdown`)

1. CLI parses args, calls `OfficeDocument.Export(path, Markdown)`.
2. `DetectKind` maps `.xlsx` to `DocumentKind.Workbook`.
3. `XlsxReader.Read` opens the package, materializes shared strings, builds a
   sparse `WorkbookModel`.
4. `WorkbookExporter.ToMarkdown` renders pipe tables per sheet.
5. CLI writes the string to stdout (or `--output` file).

## Extension points

- **New input/output format:** implement `IDocumentReader<TModel>` /
  `IDocumentWriter<TModel>` against an existing model (e.g. a CSV reader
  producing `WorkbookModel`), then surface it in `OfficeDocument`.
- **New export format:** add an `ExportFormat` member and a branch in each
  exporter's `Export` switch.
- **New template target:** `TemplateFiller.FillText` is public; a `Fill`
  overload for a new model type only needs to decide which strings to run it
  over.

## Known limitations

- No styles, number formats, merged cells, column widths, images, charts,
  headers/footers - the models are content-only.
- Formula cells are read as the formula text only; the cached computed value
  in the file is discarded, and nothing evaluates formulas.
- `.docx` reading flattens structure: tables, lists (numbering), hyperlinks
  and images are not represented; `ParagraphKind.ListItem` exists in the
  model and markdown export but the reader never produces it.
- `.pptx` is read-only (no `PptxWriter`, no `SavePresentation`) and text-only;
  slide notes are not read.
- `OfficeDocument.DetectKind` trusts the file extension, not the content.
- Template filling in documents merges a modified paragraph into one run,
  losing per-run formatting inside that paragraph.
- Everything is synchronous and loads whole documents into memory; fine for
  typical documents, not for multi-hundred-MB workbooks.
