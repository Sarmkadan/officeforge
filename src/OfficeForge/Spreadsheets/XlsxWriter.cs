using System.Globalization;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeForge.Models;

namespace OfficeForge.Spreadsheets;

public sealed class XlsxWriter : IDocumentWriter<WorkbookModel>
{
    /// <inheritdoc />
    public void Write(WorkbookModel model, Stream stream)
    {
        ArgumentNullException.ThrowIfNull(model);
        ArgumentNullException.ThrowIfNull(stream);
        using var document = SpreadsheetDocument.Create(stream, SpreadsheetDocumentType.Workbook);
        var workbookPart = document.AddWorkbookPart();
        workbookPart.Workbook = new Workbook();
        var sheets = workbookPart.Workbook.AppendChild(new Sheets());
        var sheetId = 1u;
        foreach (var sheetModel in model.Sheets.Count > 0 ? model.Sheets : [new SheetModel("Sheet1")])
        {
            var worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
            var sheetData = new SheetData();
            foreach (var rowGroup in sheetModel.OrderedCells().GroupBy(kv => kv.Key.Row))
            {
                var row = new Row { RowIndex = (uint)rowGroup.Key };
                foreach (var (cellRef, value) in rowGroup)
                    row.Append(CreateCell(cellRef, value));
                sheetData.Append(row);
            }
            worksheetPart.Worksheet = new Worksheet(sheetData);
            sheets.Append(new Sheet
            {
                Id = workbookPart.GetIdOfPart(worksheetPart),
                SheetId = sheetId++,
                Name = sheetModel.Name
            });
        }
        workbookPart.Workbook.Save();
    }

    /// <inheritdoc />
    public void Write(WorkbookModel model, string path)
    {
        ArgumentNullException.ThrowIfNull(model);
        ArgumentNullException.ThrowIfNull(path);
        ArgumentException.ThrowIfNullOrEmpty(path);
        using var stream = File.Create(path);
        Write(model, stream);
    }

    /// <summary>
    /// Asynchronously writes the workbook model to the provided <paramref name="stream"/>.
    /// </summary>
    /// <param name="model">The workbook model to write.</param>
    /// <param name="stream">The destination stream.</param>
    /// <param name="cancellationToken">Optional cancellation token.</param>
    /// <returns>A task that represents the asynchronous write operation.</returns>
    public async Task WriteAsync(WorkbookModel model, Stream stream, CancellationToken cancellationToken = default)
    {
        // The underlying OpenXML SDK does not provide async APIs, so we offload the synchronous
        // operation to a background thread to avoid blocking the calling thread.
        await Task.Run(() => Write(model, stream), cancellationToken).ConfigureAwait(false);
    }

    /// <summary>
    /// Asynchronously writes the workbook model to a file at the specified <paramref name="path"/>.
    /// </summary>
    /// <param name="model">The workbook model to write.</param>
    /// <param name="path">The file path where the workbook will be saved.</param>
    /// <param name="cancellationToken">Optional cancellation token.</param>
    /// <returns>A task that represents the asynchronous write operation.</returns>
    public async Task WriteAsync(WorkbookModel model, string path, CancellationToken cancellationToken = default)
    {
        ArgumentNullException.ThrowIfNull(model);
        ArgumentNullException.ThrowIfNull(path);
        ArgumentException.ThrowIfNullOrEmpty(path);

        // Open the file stream with async support.
        await using var stream = new FileStream(
            path,
            FileMode.Create,
            FileAccess.Write,
            FileShare.None,
            bufferSize: 4096,
            useAsync: true);

        await WriteAsync(model, stream, cancellationToken).ConfigureAwait(false);
    }

    private static Cell CreateCell(CellRef cellRef, Models.CellValue value)
    {
        var cell = new Cell { CellReference = cellRef.ToA1() };
        switch (value.Kind)
        {
            case CellValueKind.Number:
                cell.DataType = CellValues.Number;
                cell.CellValue = new DocumentFormat.OpenXml.Spreadsheet.CellValue(value.Number.ToString(CultureInfo.InvariantCulture));
                break;
            case CellValueKind.Boolean:
                cell.DataType = CellValues.Boolean;
                cell.CellValue = new DocumentFormat.OpenXml.Spreadsheet.CellValue(value.Boolean ? "1" : "0");
                break;
            case CellValueKind.DateTime:
                cell.DataType = CellValues.Date;
                cell.CellValue = new DocumentFormat.OpenXml.Spreadsheet.CellValue(value.DateTime.ToString("s", CultureInfo.InvariantCulture));
                break;
            case CellValueKind.Formula:
                cell.CellFormula = new CellFormula(value.Formula ?? string.Empty);
                break;
            default:
                cell.DataType = CellValues.InlineString;
                cell.InlineString = new InlineString(new Text(value.Text ?? string.Empty));
                break;
        }
        return cell;
    }
}
