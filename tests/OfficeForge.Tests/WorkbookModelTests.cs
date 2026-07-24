using Xunit;
using OfficeForge.Models;

namespace OfficeForge.Tests;

public class WorkbookModelTests
{
    [Fact]
    public void NewWorkbookModel_HasEmptySheetsList()
    {
        var workbook = new WorkbookModel();
        Assert.Empty(workbook.Sheets);
    }

    [Fact]
    public void AddSheet_AddsNewSheetToSheetsList()
    {
        var workbook = new WorkbookModel();
        var sheet = workbook.AddSheet("Sheet1");
        Assert.Single(workbook.Sheets);
        Assert.Same(sheet, workbook.Sheets[0]);
    }

    [Fact]
    public void FindSheet_SheetExists_ReturnsSheet()
    {
        var workbook = new WorkbookModel();
        var sheet = workbook.AddSheet("Sheet1");
        var foundSheet = workbook.FindSheet("Sheet1");
        Assert.Same(sheet, foundSheet);
    }

    [Fact]
    public void FindSheet_SheetDoesNotExist_ReturnsNull()
    {
        var workbook = new WorkbookModel();
        var foundSheet = workbook.FindSheet("Sheet1");
        Assert.Null(foundSheet);
    }

    [Fact]
    public void AddSheet_NullName_ThrowsArgumentNullException()
    {
        var workbook = new WorkbookModel();
        Assert.Throws<ArgumentNullException>(() => workbook.AddSheet(null!));
    }

    [Fact]
    public void OrderedCells_EmptyCells_ReturnsEmptyEnumerable()
    {
        var workbook = new WorkbookModel();
        var sheet = workbook.AddSheet("Sheet1");
        var orderedCells = sheet.OrderedCells();
        Assert.Empty(orderedCells);
    }
}
