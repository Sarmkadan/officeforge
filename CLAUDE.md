# CLAUDE.md

OfficeForge: cross-platform .NET 10 library + CLI to read, write, template-fill and export .xlsx/.docx/.pptx via the OpenXML SDK, without Microsoft Office or COM.

## Build

```bash
dotnet build                      # solution: OfficeForge.slnx
dotnet build -c Release           # what CI runs
dotnet pack -c Release --no-build -o artifacts   # NuGet packages (OfficeForge, OfficeForge.Cli)
dotnet run --project src/OfficeForge.Cli -- read-cell file.xlsx B2
```

## Test

```bash
dotnet test                                       # all xUnit tests
dotnet test -c Release --no-build                 # CI form (after Release build)
dotnet test --filter "FullyQualifiedName~CellRefTests"
```

CI: `.github/workflows/ci.yml` (build -> test -> pack on push/PR, .NET 10.0.x).

## Lint / format

No analyzers or `.editorconfig` configured. `Nullable` and `ImplicitUsings` are enabled in every project; treat nullable warnings as errors to fix. Use `dotnet format` before committing if you touch formatting.

## Layout

- `OfficeForge.slnx` - solution (XML slnx format)
- `src/OfficeForge/` - the library (NuGet `OfficeForge`)
  - `OfficeDocument.cs` - static facade and main entry point (`OpenWorkbook`, `SaveWorkbook`, `Export`, `DetectKind`)
  - `Models/` - format-agnostic in-memory models; must not reference OpenXML (`WorkbookModel`, `SheetModel`, `CellRef`, `CellValue`, `DocumentModel`, `PresentationModel`)
  - `Spreadsheets/` `XlsxReader`/`XlsxWriter`; `Words/` `DocxReader`/`DocxWriter`; `Slides/` `PptxReader`/`PptxWriter`
  - `Export/` - `IDocumentExporter` + text/markdown/json exporters
  - `Templates/` - `TemplateFiller` (`{{key}}` replacement in models)
  - `IDocumentReader.cs` - `IDocumentReader<T>` / `IDocumentWriter<T>` with Stream and path overloads
  - `ReaderOptions.cs`, `DocumentTooLargeException.cs`, `OfficeForgeFormatException.cs`
- `src/OfficeForge.Cli/Program.cs` - top-level-statements CLI (System.CommandLine), tool command `officeforge`
- `tests/OfficeForge.Tests/` - xUnit 2.9 tests
- `docs/ARCHITECTURE.md` - module breakdown, design decisions, limitations (read before structural changes)
- `src/OfficeForge.Tests/*.cs` and `test_zip_protection.csx` - stray files not in any csproj; not compiled

## Conventions

- Readers/writers/exporters/filler are stateless; only models hold mutable state. OpenXML types never appear in the public API.
- Names: `{Format}Reader` / `{Format}Writer` (`XlsxReader`, `DocxWriter`), `*Model` for data types, `I*` interfaces.
- `CellValue` built only via `CellValue.From*` factories; `CellRef` is 1-based `(Row, Column)` with A1 parsing (`Parse`/`TryParse`/`ToA1`).
- Missing cell reads return `CellValue.Empty`, never throw. Sheets are sparse dictionaries; iterate via `OrderedCells()`.
- Invariant culture for all number/date formatting and parsing.
- Public API carries XML doc comments; guard args with `ArgumentException.ThrowIfNullOrWhiteSpace` etc.
- Tests: `MethodOrScenario_Condition_Expected` naming, Arrange/Act/Assert comments, `[Fact]`/`[Theory]`; fixtures that create temp files implement `IDisposable`.
- Commits: conventional prefixes (`feat:`, `fix(scope):`, `docs:`, `chore:`).
