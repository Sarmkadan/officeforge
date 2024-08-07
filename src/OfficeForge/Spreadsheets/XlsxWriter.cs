using System.Globalization;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeForge.Models;

namespace OfficeForge.Spreadsheets;

public sealed class XlsxWriter : IDocumentWriter<WorkbookModel>
{
    public void Write(WorkbookModel model, string path)
    {
        using var stream = File.Create(path);
        Write(model, stream);
    }

    public void Write(WorkbookModel model, Stream stream)
    {
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
