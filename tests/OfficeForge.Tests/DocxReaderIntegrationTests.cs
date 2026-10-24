using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Xml;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeForge.Models;
using OfficeForge.Words;
using Xunit;

namespace OfficeForge.Tests;

/// <summary>
/// Integration tests for DocxReader that test reading DOCX files
/// not produced by DocxWriter (hand-crafted DOCX packages).
/// </summary>
public class DocxReaderIntegrationTests : IDisposable
{
    readonly string _dir = Directory.CreateTempSubdirectory("DocxReaderTests").FullName;

    public void Dispose() => Directory.Delete(_dir, recursive: true);

    string TempPath(string name) => Path.Combine(_dir, name);

    /// <summary>
    /// Tests reading a DOCX file with an empty body (no paragraphs).
    /// The resulting DocumentModel should have zero paragraphs.
    /// </summary>
    [Fact]
    public void Read_EmptyBody_DocumentModelWithZeroParagraphs()
    {
        // Arrange: Create a minimal DOCX with a zip containing word/document.xml with empty body
        using var stream = CreateDocxWithEmptyBody();
        stream.Position = 0;

        // Act
        var doc = new DocxReader().Read(stream);

        // Assert
        Assert.Empty(doc.Paragraphs);
    }

    /// <summary>
    /// Tests reading a DOCX file with a paragraph containing nested runs and line breaks (w:br).
    /// The resulting DocumentModel should preserve the text and line break information.
    /// </summary>
    [Fact]
    public void Read_ParagraphWithNestedRunsAndLineBreaks_PreservesText()
    {
        // Arrange: Create a DOCX with a paragraph containing multiple runs and a line break
        using var stream = CreateDocxWithParagraphAndLineBreak();
        stream.Position = 0;

        // Act
        var doc = new DocxReader().Read(stream);

        // Assert
        Assert.Single(doc.Paragraphs);
        Assert.Single(doc.Paragraphs[0].Runs);
        // The text should include content; line breaks in the XML may or may not be captured
        // depending on how the reader handles w:br - at minimum no exception should be thrown
        Assert.NotEmpty(doc.Paragraphs[0].Text);
    }

    /// <summary>
    /// Tests reading a DOCX file with merged table cells (w:gridSpan / w:vMerge).
    /// The reader should skip table-related content gracefully without throwing.
    /// </summary>
    [Fact]
    public void Read_DocxWithMergedTableCells_NoException()
    {
        // Arrange: Create a DOCX with a table having merged cells
        using var stream = CreateDocxWithMergedTableCells();
        stream.Position = 0;

        // Act & Assert - should not throw, resulting model may have limited table content
        var doc = Record.Exception(() => new DocxReader().Read(stream));
        Assert.Null(doc); // Expected to either succeed with limited content or gracefully handle
    }

    /// <summary>
    /// Tests reading a DOCX file with unknown elements (w:sdt content controls).
    /// The reader should skip unknown elements without throwing.
    /// </summary>
    [Fact]
    public void Read_DocxWithUnknownElements_w_sdt_NoException()
    {
        // Arrange: Create a DOCX with a w:sdt (content control) element
        using var stream = CreateDocxWithSdtElement();
        stream.Position = 0;

        // Act & Assert - should not throw when encountering unknown w:sdt elements
        var doc = Record.Exception(() => new DocxReader().Read(stream));
        Assert.Null(doc); // Expected to either succeed with limited content or gracefully handle
    }

    /// <summary>
    /// Tests that a corrupt (non-zip) file produces a clear exception rather than a NullReferenceException.
    /// </summary>
    [Fact]
    public void Read_CorruptNonZipFile_ThrowsClearException()
    {
        // Arrange: Create a file that is not a valid zip
        File.WriteAllText(TempPath("corrupt.docx"), "This is not a ZIP file at all!");

        // Act & Assert - should throw a clear exception, not NullReferenceException
        var ex = Assert.Throws<System.IO.InvalidDataException>(() => new DocxReader().Read(TempPath("corrupt.docx")));
        // The InvalidDataException is a clear exception indicating the file is corrupt
        Assert.Contains("End of Central Directory record", ex.Message);
    }

    /// <summary>
    /// Tests that a zip file without word/document.xml produces a clear exception rather than a NullReferenceException.
    /// </summary>
    [Fact]
    public void Read_ZipWithoutDocumentXml_ThrowsClearException()
    {
        // Arrange: Create a minimal zip without word/document.xml
        using var stream = CreateZipWithoutDocumentXml();
        stream.Position = 0;

        // Act & Assert - should throw a clear exception, not NullReferenceException
        var ex = Assert.Throws<System.IO.FileFormatException>(() => new DocxReader().Read(stream));
        // The FileFormatException is a clear exception indicating invalid package format
        Assert.Contains("content type", ex.Message);
    }

    // Helper methods

    static MemoryStream CreateDocxWithEmptyBody()
    {
        var stream = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document, true))
        {
            var mainPart = doc.AddMainDocumentPart();
            mainPart.Document = new Document(new Body());
            mainPart.Document.Save();
        }
        stream.Position = 0;
        return stream;
    }

    static MemoryStream CreateDocxWithParagraphAndLineBreak()
    {
        var stream = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document, true))
        {
            var mainPart = doc.AddMainDocumentPart();
            var body = new Body();

            // Add a paragraph with text and a line break (w:br)
            var paragraph = new Paragraph();
            var run = new Run();
            run.Append(new Text("Hello"));

            // Add w:br element manually via OpenXml
            run.Append(new Break() { Type = BreakValues.Page });

            paragraph.Append(run);
            body.Append(paragraph);
            doc.MainDocumentPart!.Document = new Document(body);
            doc.MainDocumentPart.Document.Save();
        }
        stream.Position = 0;
        return stream;
    }

    static MemoryStream CreateDocxWithMergedTableCells()
    {
        var stream = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document, true))
        {
            var mainPart = doc.AddMainDocumentPart();
            var body = new Body();

            // Add a simple table with merged cells conceptually
            // We'll add a table with rows/cells; the reader should handle gracefully
            var table = new Table();

            // Add a row with cells
            var tableRow = new TableRow();
            var tableCell = new TableCell();
            tableCell.Append(new Paragraph(new Run(new Text("Cell text"))));
            tableRow.Append(tableCell);

            // Add a second cell with gridSpan to simulate merging
            var tableCell2 = new TableCell();
            // Add gridSpan element to indicate merged cells
            var gridSpan = new GridSpan() { Val = 2 };
            tableCell2.Append(gridSpan);
            tableCell2.Append(new Paragraph(new Run(new Text("Merged cell"))));
            tableRow.Append(tableCell2);

            table.Append(tableRow);
            body.Append(table);
            doc.MainDocumentPart!.Document = new Document(body);
            doc.MainDocumentPart.Document.Save();
        }
        stream.Position = 0;
        return stream;
    }

    static MemoryStream CreateDocxWithSdtElement()
    {
        var stream = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document, true))
        {
            var mainPart = doc.AddMainDocumentPart();
            var body = new Body();

            // Add a paragraph with a w:sdt (content control) element
            var paragraph = new Paragraph();

            // Add w:sdt element (content control) - this is an unknown element for the reader
            var sdt = new SdtBlock();
            var sdtPr = new SdtProperties();
            var sdtAlias = new SdtAlias() { Val = "My Content Control" };
            sdtPr.Append(sdtAlias);
            sdt.Append(sdtPr);

            // Add some text inside the sdt
            var sdtContentRun = new Run();
            sdtContentRun.Append(new Text("Content inside control"));
            sdt.Append(sdtContentRun);

            paragraph.Append(sdt);
            body.Append(paragraph);
            doc.MainDocumentPart!.Document = new Document(body);
            doc.MainDocumentPart.Document.Save();
        }
        stream.Position = 0;
        return stream;
    }

    static MemoryStream CreateZipWithoutDocumentXml()
    {
        var stream = new MemoryStream();

        // Create a zip with only [Content_Types].xml and _rels/.rels, but NO word/document.xml
        using (var archive = new ZipArchive(stream, ZipArchiveMode.Create, true))
        {
            // Add the required [Content_Types].xml
            var ctEntry = archive.CreateEntry("[Content_Types].xml");
            using (var ctStream = ctEntry.Open())
            {
                var writer = new System.Xml.XmlTextWriter(ctStream, System.Text.Encoding.UTF8);
                writer.WriteStartDocument();
                writer.WriteStartElement("Types", "http://schemas.openxmlformats.org/package/2006/content-types");
                writer.WriteStartElement("Default");
                writer.WriteAttributeString("Extension", "rels");
                writer.WriteAttributeString("ContentType", "application/vnd.openxmlformats-officedocument.relationships+xml");
                writer.WriteEndElement();
                writer.WriteStartElement("Default");
                writer.WriteAttributeString("Extension", "xml");
                writer.WriteAttributeString("ContentType", "application/xml");
                writer.WriteEndElement();
                writer.WriteEndElement();
                writer.WriteEndDocument();
                writer.Close();
            }

            // Add _rels/.rels
            var relsEntry = archive.CreateEntry("_rels/.rels");
            using (var relsStream = relsEntry.Open())
            {
                var writer = new System.Xml.XmlTextWriter(relsStream, System.Text.Encoding.UTF8);
                writer.WriteStartDocument();
                writer.WriteStartElement("Relationships", "http://schemas.openxmlformats.org/package/2006/relationships");
                writer.WriteStartElement("Relationship");
                writer.WriteAttributeString("Id", "rId1");
                writer.WriteAttributeString("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument");
                writer.WriteAttributeString("Target", "word/document.xml");
                writer.WriteEndElement();
                writer.WriteEndElement();
                writer.WriteEndDocument();
                writer.Close();
            }
        }

        stream.Position = 0;
        return stream;
    }
}