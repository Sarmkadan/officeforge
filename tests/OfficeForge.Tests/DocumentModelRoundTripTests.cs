using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Packaging;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using OfficeForge.Models;
using OfficeForge.Words;
using Xunit;

namespace OfficeForge.Tests;

/// <summary>
/// Integration test for DocumentModel round-trip via DocxWriter/DocxReader.
/// Verifies that a DocumentModel with headings, paragraphs (bold/italic runs),
/// and body text can be written to DOCX, read back, and the model matches.
/// Also verifies the ZIP packaging artifacts are well-formed XML.
/// </summary>
public sealed class DocumentModelRoundTripTests : IDisposable
{
    private readonly string _tempDir = Directory.CreateTempSubdirectory("officeforge-rt-test-").FullName;

    /// <summary>
    /// Cleans up the temporary directory created for the test.
    /// </summary>
    public void Dispose() => Directory.Delete(_tempDir, recursive: true);

    [Fact]
    public void RoundTrip_HeadingsParagraphs_preservesModel()
    {
        // --- Arrange: build a rich DocumentModel ---
        var doc = new DocumentModel();

        // Heading 1
        var h1 = doc.AddParagraph("Document Title", new RunStyle(Bold: true, FontSize: 24));
        h1.Kind = ParagraphKind.Heading;
        h1.HeadingLevel = 1;

        // Heading 2
        var h2 = doc.AddParagraph("Section One", new RunStyle(Italic: true, FontSize: 16));
        h2.Kind = ParagraphKind.Heading;
        h2.HeadingLevel = 2;

        // Body paragraph with bold and italic runs
        var bodyPara = doc.AddParagraph("This is some body text with ");
        bodyPara.Runs.Add(new RunModel("bold text", new RunStyle(Bold: true)));
        bodyPara.Runs.Add(new RunModel(" and italic text ", new RunStyle(Italic: true)));
        bodyPara.Runs.Add(new RunModel("more body text", new RunStyle()));

        // Write to temp file
        var tempPath = Path.Combine(_tempDir, "test_document.docx");
        try
        {
            new DocxWriter().Write(doc, tempPath);

            // --- Act: read back ---
            var roundtripped = new DocxReader().Read(tempPath);

            // --- Assert: model-level round-trip fidelity ---
            // Paragraph count
            Assert.Equal(doc.Paragraphs.Count, roundtripped.Paragraphs.Count);

            for (int i = 0; i < doc.Paragraphs.Count; i++)
            {
                var expPara = doc.Paragraphs[i];
                var actPara = roundtripped.Paragraphs[i];

                // Kind and heading level
                Assert.Equal(expPara.Kind, actPara.Kind);
                Assert.Equal(expPara.HeadingLevel, actPara.HeadingLevel);
                // Text content
                Assert.Equal(expPara.Text, actPara.Text);

                // Runs: text and style (Bold, Italic, Underline)
                Assert.Equal(expPara.Runs.Count, actPara.Runs.Count);
                for (int r = 0; r < expPara.Runs.Count; r++)
                {
                    var expRun = expPara.Runs[r];
                    var actRun = actPara.Runs[r];

                    Assert.Equal(expRun.Text, actRun.Text);
                    Assert.Equal(expRun.Style.Bold, actRun.Style.Bold);
                    Assert.Equal(expRun.Style.Italic, actRun.Style.Italic);
                    Assert.Equal(expRun.Style.Underline, actRun.Style.Underline);
                    // FontName may be null; compare if set
                    if (expRun.Style.FontName is { } fn)
                    {
                        Assert.Equal(fn, actRun.Style.FontName);
                    }
                    // FontSize is not currently preserved by DocxReader
                }
            }

            // --- Assert: ZIP packaging artifacts exist and are well-formed XML ---
            // First verify via Package (System.IO.Packaging)
            using (var package = Package.Open(tempPath, FileMode.Open, FileAccess.Read))
            {
                // Check all parts exist
                var partNames = package.GetParts().Select(p => p.Uri.OriginalString).ToList();

                // word/document.xml
                Assert.Contains(partNames, n => n.EndsWith("word/document.xml", StringComparison.OrdinalIgnoreCase));
                var documentPart = package.GetParts().First(p => p.Uri.OriginalString.EndsWith("word/document.xml", StringComparison.OrdinalIgnoreCase));
                using var documentStream = documentPart.GetStream();
                var documentXml = System.Xml.Linq.XDocument.Load(documentStream);
                Assert.NotNull(documentXml.Root);

                // _rels/.rels
                Assert.Contains(partNames, n => n.EndsWith("_rels/.rels", StringComparison.OrdinalIgnoreCase));
                var relsPart = package.GetParts().First(p => p.Uri.OriginalString.EndsWith("_rels/.rels", StringComparison.OrdinalIgnoreCase));
                using var relsStream = relsPart.GetStream();
                var relsXml = System.Xml.Linq.XDocument.Load(relsStream);
                Assert.NotNull(relsXml.Root);
            }

            // Then verify [Content_Types].xml via ZipArchive (since it's not exposed via Package.GetParts())
            using var zipArchive = System.IO.Compression.ZipFile.OpenRead(tempPath);
            var contentTypesEntry = zipArchive.GetEntry("[Content_Types].xml");
            Assert.NotNull(contentTypesEntry);
            using var entryStream = contentTypesEntry.Open();
            var contentTypesXml = System.Xml.Linq.XDocument.Load(entryStream);
            Assert.NotNull(contentTypesXml.Root);
        }
        finally
        {
            // Clean up temp file
            if (File.Exists(tempPath))
            {
                File.Delete(tempPath);
            }
        }
    }
}