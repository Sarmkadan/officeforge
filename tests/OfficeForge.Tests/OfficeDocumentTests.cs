using Xunit;
using OfficeForge;
using OfficeForge.Models;
using OfficeForge.Export;

namespace OfficeForge.Tests;

public class OfficeDocumentTests
{
    [Theory]
    [InlineData("test.xlsx", DocumentKind.Workbook)]
    [InlineData("test.xlsm", DocumentKind.Workbook)]
    [InlineData("test.docx", DocumentKind.Document)]
    [InlineData("test.pptx", DocumentKind.Presentation)]
    public void DetectKind_ReturnsCorrectKind_ForSupportedExtensions(string fileName, DocumentKind expectedKind)
    {
        var result = OfficeDocument.DetectKind(fileName);
        Assert.Equal(expectedKind, result);
    }

    [Theory]
    [InlineData(null)]
    [InlineData("")]
    [InlineData("   ")]
    public void DetectKind_ThrowsArgumentException_ForNullOrWhitespace(string fileName)
    {
        Assert.Throws<ArgumentException>(() => OfficeDocument.DetectKind(fileName));
    }

    [Fact]
    public void DetectKind_ThrowsNotSupportedException_ForUnsupportedExtension()
    {
        Assert.Throws<NotSupportedException>(() => OfficeDocument.DetectKind("file.txt"));
    }

    [Fact]
    public void SaveWorkbook_ThrowsArgumentNullException_ForNullWorkbook()
    {
        Assert.Throws<ArgumentNullException>(() => OfficeDocument.SaveWorkbook(null!, "output.xlsx"));
    }

    [Fact]
    public void SaveDocument_ThrowsArgumentNullException_ForNullDocument()
    {
        Assert.Throws<ArgumentNullException>(() => OfficeDocument.SaveDocument(null!, "output.docx"));
    }

    [Fact]
    public void OpenWorkbook_ThrowsException_ForNonExistentFile()
    {
        // Assuming XlsxReader throws when file does not exist
        Assert.ThrowsAny<Exception>(() => OfficeDocument.OpenWorkbook("nonexistent.xlsx"));
    }

    [Fact]
    public void Export_ThrowsException_ForNonExistentFile()
    {
        // Assuming the underlying reader throws when file does not exist
        Assert.ThrowsAny<Exception>(() => OfficeDocument.Export("nonexistent.docx", ExportFormat.Markdown));
    }
}
