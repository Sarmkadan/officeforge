using OfficeForge;
using OfficeForge.Models;
using Xunit;
using System;

namespace OfficeForge.Tests;

/// <summary>
/// Provides round-trip tests for Excel (.xlsx) files, ensuring that writing and reading preserves workbook content including cell values, formulas, multiple sheets, and default sheet behavior.
/// </summary>
public class XlsxRoundTripTests : IDisposable, IXlsxRoundTripTests, IEquatable<XlsxRoundTripTests>
{
    readonly string _dir = Directory.CreateTempSubdirectory(XlsxRoundTripTestsConstants.TempDirectoryPrefix).FullName;

    /// <summary>
    /// Releases the temporary directory used for test files.
    /// </summary>
    public void Dispose() => Directory.Delete(_dir, recursive: true);

    string TempPath(string name) => Path.Combine(_dir, name);

    // Equality members
    /// <summary>
    /// Indicates whether the current object is equal to another object of the same type.
    /// </summary>
    /// <param name="other">An object to compare with this object.</param>
    /// <returns>true if the current object is equal to the <paramref name="other">parameter</paramref>; otherwise, false.</returns>
    public bool Equals(XlsxRoundTripTests? other)
    {
        if (ReferenceEquals(null, other)) return false;
        if (ReferenceEquals(this, other)) return true;
        return _dir == other._dir;
    }

    /// <summary>
    /// Indicates whether the current object is equal to a specified object.
    /// </summary>
    /// <param name="obj">The object to compare with the current object.</param>
    /// <returns>true if the specified object  is equal to the current object; otherwise, false.</returns>
    public override bool Equals(object? obj) => Equals(obj as XlsxRoundTripTests);

    /// <summary>
    /// Serves as the default hash function.
    /// </summary>
    /// <returns>A hash code for the current object.</returns>
    public override int GetHashCode() => HashCode.Combine(_dir);

    /// <summary>
    /// Determines whether two specified objects are equal.
    /// </summary>
    /// <param name="left">The first object to compare.</param>
    /// <param name="right">The second object to compare.</param>
    /// <returns>true if the values of <paramref name="left"/> and <paramref name="right"/> are equal; otherwise, false.</returns>
    public static bool operator ==(XlsxRoundTripTests? left, XlsxRoundTripTests? right) => Equals(left, right);

    /// <summary>
    /// Determines whether two specified objects are not equal.
    /// </summary>
    /// <param name="left">The first object to compare.</param>
    /// <param name="right">The second object to compare.</param>
    /// <returns>true if the values of <paramref name="left"/> and <paramref name="right"/> are not equal; otherwise, false.</returns>
    public static bool operator !=(XlsxRoundTripTests? left, XlsxRoundTripTests? right) => !Equals(left, right);

    [Fact]
    /// <summary>
    /// Verifies that typed cell values (text, number, boolean, date/time, formula) are preserved when saving and reloading a workbook.
    /// </summary>
    public void WriteRead_PreservesTypedCellValues()
    {
        var workbook = new WorkbookModel();
        var sheet = workbook.AddSheet("Data");
        sheet["A1"] = CellValue.FromText(XlsxRoundTripTestsConstants.SampleRevenueText);
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
        Assert.Equal(XlsxRoundTripTestsConstants.SampleRevenueText, data["A1"].Text);
        Assert.Equal(1234.5, data["B1"].Number);
        Assert.True(data["C1"].Boolean);
        Assert.Equal(new DateTime(2026, 3, 15, 10, 30, 0), data["D1"].DateTime);
        Assert.Equal(CellValueKind.Formula, data["A2"].Kind);
        Assert.Equal("SUM(B1:B1)", data["A2"].Formula);
    }

    [Fact]
    /// <summary>
    /// Verifies that multiple worksheets are preserved when saving and reloading a workbook.
    /// </summary>
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
    /// <summary>
    /// Verifies that an empty workbook gets a default worksheet named "Sheet1" when saved and reloaded.
    /// </summary>
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
    /// <summary>
    /// Verifies that exporting an Excel workbook to Markdown produces a table containing the expected text and numeric values.
    /// </summary>
    public void Export_FromXlsxPath_ProducesMarkdownTable()
    {
        var workbook = new WorkbookModel();
        var sheet = workbook.AddSheet("Report");
        sheet["A1"] = CellValue.FromText(XlsxRoundTripTestsConstants.SampleRevenueText);
        sheet["B1"] = CellValue.FromNumber(42500.75);
        var path = TempPath("report.xlsx");
        OfficeDocument.SaveWorkbook(workbook, path);
        var markdown = OfficeDocument.Export(path, Export.ExportFormat.Markdown);
        Assert.Contains("## Report", markdown);
        Assert.Contains($"| {XlsxRoundTripTestsConstants.SampleRevenueText} | 42500.75 |", markdown);
    }

    [Fact]
    /// <summary>
    /// Verifies that reading a cell that does not exist returns an empty cell value.
    /// </summary>
    public void MissingCell_ReadsAsEmpty()
    {
        var sheet = new WorkbookModel().AddSheet("S");
        Assert.True(sheet["Z99"].IsEmpty);
    }
}

public class DocxRoundTripTests : IDisposable
{
    readonly string _dir = Directory.CreateTempSubdirectory(XlsxRoundTripTestsConstants.TempDirectoryPrefix).FullName;

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