using Xunit;
using OfficeForge.Models;
using OfficeForge.Export;

namespace OfficeForge.Tests;

public class WorkbookExporterTests
{
    private static WorkbookModel CreateSimpleWorkbook()
    {
        var workbook = new WorkbookModel();
        var sheet = workbook.AddSheet("Sheet1");
        sheet["A1"] = CellValue.FromText("Name");
        sheet["B1"] = CellValue.FromText("Age");
        sheet["A2"] = CellValue.FromText("Alice");
        sheet["B2"] = CellValue.FromNumber(30);
        return workbook;
    }

    [Fact]
    public void Export_Markdown_ReturnsCorrectFormat()
    {
        var workbook = CreateSimpleWorkbook();
        var result = WorkbookExporter.Export(workbook, ExportFormat.Markdown);
        Assert.Contains("## Sheet1", result);
        Assert.Contains("| Name | Age |", result);
    }

    [Fact]
    public void Export_Json_ReturnsValidJson()
    {
        var workbook = CreateSimpleWorkbook();
        var result = WorkbookExporter.Export(workbook, ExportFormat.Json);
        Assert.Contains("\"Sheet1\"", result);
        Assert.Contains("\"A1\": \"Name\"", result);
    }

    [Fact]
    public void ToPlainText_ReturnsTabSeparatedValues()
    {
        var workbook = CreateSimpleWorkbook();
        var result = WorkbookExporter.ToPlainText(workbook);
        Assert.Contains("Sheet1", result);
        Assert.Contains("Name\tAge", result);
        Assert.Contains("Alice\t30", result);
    }

    [Fact]
    public void ToMarkdown_EmptyWorkbook_ReturnsEmptyString()
    {
        var workbook = new WorkbookModel();
        var result = WorkbookExporter.ToMarkdown(workbook);
        Assert.Equal(string.Empty, result);
    }

    [Fact]
    public void ToJson_EmptyWorkbook_ReturnsEmptyObject()
    {
        var workbook = new WorkbookModel();
        var result = WorkbookExporter.ToJson(workbook);
        Assert.Equal("{}", result.Trim());
    }

    [Fact]
    public void ToPlainText_NullWorkbook_ThrowsNullReferenceException()
    {
        Assert.Throws<NullReferenceException>(() => WorkbookExporter.ToPlainText(null!));
    }
}
