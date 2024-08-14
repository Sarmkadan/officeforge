using OfficeForge;
using OfficeForge.Models;
using Xunit;

namespace OfficeForge.Tests;

public class XlsxRoundTripTests : IDisposable
{
    readonly string _dir = Directory.CreateTempSubdirectory("officeforge-tests-").FullName;

    public void Dispose() => Directory.Delete(_dir, recursive: true);

    string TempPath(string name) => Path.Combine(_dir, name);

    [Fact]
    public void WriteRead_PreservesTypedCellValues()
    {
        var workbook = new WorkbookModel();
        var sheet = workbook.AddSheet("Data");
        sheet["A1"] = CellValue.FromText("Revenue");
        sheet["B1"] = CellValue.FromNumber(1234.5);
        sheet["C1"] = CellValue.FromBoolean(true);
        sheet["D1"] = CellValue.FromDateTime(new DateTime(2026, 3, 15, 10, 30, 0));
        sheet["A2"] = CellValue.FromFormula("SUM(B1:B1)");
        var path = TempPath("typed.xlsx");
        OfficeDocument.SaveWorkbook(workbook, path);
        var loaded = OfficeDocument.OpenWorkbook(path);
        var data = Assert.Single(loaded.Sheets);
        Assert.Equal("Data", data.Name);
        Assert.Equal(CellValueKind.Text, data["A1"].Kind);
        Assert.Equal("Revenue", data["A1"].Text);
        Assert.Equal(1234.5, data["B1"].Number);
        Assert.True(data["C1"].Boolean);
        Assert.Equal(new DateTime(2026, 3, 15, 10, 30, 0), data["D1"].DateTime);
        Assert.Equal(CellValueKind.Formula, data["A2"].Kind);
        Assert.Equal("SUM(B1:B1)", data["A2"].Formula);
    }

    [Fact]
    public void WriteRead_PreservesMultipleSheets()
    {
        var workbook = new WorkbookModel();
        workbook.AddSheet("First")["A1"] = CellValue.FromNumber(1);
        workbook.AddSheet("Second")["B2"] = CellValue.FromNumber(2);
        var path = TempPath("multi.xlsx");
        OfficeDocument.SaveWorkbook(workbook, path);
        var loaded = OfficeDocument.OpenWorkbook(path);
        Assert.Equal(["First", "Second"], loaded.Sheets.Select(s => s.Name));
        Assert.Equal(2, loaded.FindSheet("second")!["B2"].Number);
    }

    [Fact]
    public void WriteRead_EmptyWorkbookGetsDefaultSheet()
    {
        var path = TempPath("empty.xlsx");
        OfficeDocument.SaveWorkbook(new WorkbookModel(), path);
        var loaded = OfficeDocument.OpenWorkbook(path);
        var sheet = Assert.Single(loaded.Sheets);
        Assert.Equal("Sheet1", sheet.Name);
        Assert.Empty(sheet.Cells);
    }

    [Fact]
    public void Export_FromXlsxPath_ProducesMarkdownTable()
    {
        var workbook = new WorkbookModel();
        var sheet = workbook.AddSheet("Report");
        sheet["A1"] = CellValue.FromText("Revenue");
        sheet["B1"] = CellValue.FromNumber(42500.75);
        var path = TempPath("report.xlsx");
        OfficeDocument.SaveWorkbook(workbook, path);
        var markdown = OfficeDocument.Export(path, Export.ExportFormat.Markdown);
        Assert.Contains("## Report", markdown);
        Assert.Contains("| Revenue | 42500.75 |", markdown);
    }

    [Fact]
    public void MissingCell_ReadsAsEmpty()
    {
        var sheet = new WorkbookModel().AddSheet("S");
        Assert.True(sheet["Z99"].IsEmpty);
    }
}

public class DocxRoundTripTests : IDisposable
{
    readonly string _dir = Directory.CreateTempSubdirectory("officeforge-tests-").FullName;

    public void Dispose() => Directory.Delete(_dir, recursive: true);

    [Fact]
    public void WriteRead_PreservesParagraphsAndStyles()
    {
        var doc = new DocumentModel();
        var heading = doc.AddParagraph("Quarterly Report");
        heading.Kind = ParagraphKind.Heading;
        heading.HeadingLevel = 1;
        doc.AddParagraph("Plain body text.");
        doc.AddParagraph("Bold statement", new RunStyle(Bold: true));
        var path = Path.Combine(_dir, "report.docx");
        OfficeDocument.SaveDocument(doc, path);
        var loaded = OfficeDocument.OpenDocument(path);
        Assert.Equal(3, loaded.Paragraphs.Count);
        Assert.Equal(ParagraphKind.Heading, loaded.Paragraphs[0].Kind);
        Assert.Equal(1, loaded.Paragraphs[0].HeadingLevel);
        Assert.Equal("Quarterly Report", loaded.Paragraphs[0].Text);
        Assert.Equal("Plain body text.", loaded.Paragraphs[1].Text);
        var bold = loaded.Paragraphs[2].Runs.Single();
        Assert.True(bold.Style.Bold);
        Assert.Equal("Bold statement", bold.Text);
    }

    [Fact]
    public void ToPlainText_JoinsParagraphs()
    {
        var doc = new DocumentModel();
        doc.AddParagraph("one");
        doc.AddParagraph("two");
        Assert.Equal("one" + Environment.NewLine + "two", doc.ToPlainText());
    }

    [Fact]
    public void DetectKind_MapsExtensions()
    {
        Assert.Equal(DocumentKind.Workbook, OfficeDocument.DetectKind("a.xlsx"));
        Assert.Equal(DocumentKind.Document, OfficeDocument.DetectKind("b.DOCX"));
        Assert.Equal(DocumentKind.Presentation, OfficeDocument.DetectKind("c.pptx"));
        Assert.Throws<NotSupportedException>(() => OfficeDocument.DetectKind("d.pdf"));
    }
}
