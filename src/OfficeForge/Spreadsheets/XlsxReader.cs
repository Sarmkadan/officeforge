using System.Globalization;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeForge.Models;

namespace OfficeForge.Spreadsheets;

public sealed class XlsxReader : IDocumentReader<WorkbookModel>
{
    public WorkbookModel Read(string path)
    {
        using var stream = File.OpenRead(path);
        return Read(stream);
    }

    public WorkbookModel Read(Stream stream)
    {
        using var document = SpreadsheetDocument.Open(stream, false);
        var workbookPart = document.WorkbookPart ?? throw new InvalidDataException("Workbook part missing.");
        var sharedStrings = workbookPart.SharedStringTablePart?.SharedStringTable?
            .Elements<SharedStringItem>().Select(i => i.InnerText).ToArray() ?? [];
        var model = new WorkbookModel();
        foreach (var sheet in workbookPart.Workbook.Descendants<Sheet>())
        {
            if (sheet.Id?.Value is not { } relId || sheet.Name?.Value is not { } name) continue;
            if (workbookPart.GetPartById(relId) is not WorksheetPart worksheetPart) continue;
            var sheetModel = model.AddSheet(name);
            foreach (var cell in worksheetPart.Worksheet.Descendants<Cell>())
            {
                if (cell.CellReference?.Value is not { } reference) continue;
                if (!CellRef.TryParse(reference, out var cellRef)) continue;
                var value = ConvertCell(cell, sharedStrings);
                if (!value.IsEmpty) sheetModel[cellRef] = value;
            }
        }
        return model;
    }

    private static Models.CellValue ConvertCell(Cell cell, string[] sharedStrings)
    {
        if (cell.CellFormula?.Text is { Length: > 0 } formula)
            return Models.CellValue.FromFormula(formula);
        var raw = cell.CellValue?.InnerText;
        if (string.IsNullOrEmpty(raw) && cell.DataType?.Value != CellValues.InlineString)
            return Models.CellValue.Empty;
        var type = cell.DataType?.Value;
        if (type == CellValues.SharedString)
            return int.TryParse(raw, out var index) && index >= 0 && index < sharedStrings.Length
                ? Models.CellValue.FromText(sharedStrings[index])
                : Models.CellValue.Empty;
        if (type == CellValues.Boolean)
            return Models.CellValue.FromBoolean(raw == "1");
        if (type == CellValues.String || type == CellValues.InlineString)
            return Models.CellValue.FromText(type == CellValues.InlineString ? cell.InnerText : raw ?? string.Empty);
        if (type == CellValues.Date && DateTime.TryParse(raw, CultureInfo.InvariantCulture, DateTimeStyles.RoundtripKind, out var date))
            return Models.CellValue.FromDateTime(date);
        return double.TryParse(raw, NumberStyles.Float, CultureInfo.InvariantCulture, out var number)
            ? Models.CellValue.FromNumber(number)
            : Models.CellValue.FromText(raw ?? string.Empty);
    }
}
