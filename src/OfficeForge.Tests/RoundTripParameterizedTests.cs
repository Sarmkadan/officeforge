using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using OfficeForge;
using OfficeForge.Models;
using Xunit;

namespace OfficeForge.Tests;

/// <summary>
/// Provides a parameterized round‑trip test suite that verifies that reading a
/// document, writing it back, and reading it again yields a model equivalent to
/// the original one. The suite runs over a small corpus of fixture files for
/// each supported format (XLSX, DOCX, PPTX).
/// </summary>
public sealed class RoundTripParameterizedTests : IDisposable
{
    private readonly string _tempDir = Directory.CreateTempSubdirectory("officeforge-rt-").FullName;

    /// <summary>
    /// Cleans up the temporary directory created for the tests.
    /// </summary>
    public void Dispose() => Directory.Delete(_tempDir, recursive: true);

    /// <summary>
    /// Returns the full paths of all fixture files located in the
    /// <c>Tests/OfficeForge.Tests/Fixtures</c> directory.
    /// </summary>
    public static IEnumerable<object[]> FixtureFiles()
    {
        var baseDir = Path.GetDirectoryName(typeof(RoundTripParameterizedTests).Assembly.Location)
            ?? throw new InvalidOperationException("Unable to locate assembly directory.");

        var fixturesDir = Path.Combine(baseDir, "Fixtures");
        if (!Directory.Exists(fixturesDir))
            yield break; // No fixtures – the test will be skipped.

        var supportedExtensions = new[] { ".xlsx", ".docx", ".pptx" };
        foreach (var file in Directory.EnumerateFiles(fixturesDir, "*.*", SearchOption.AllDirectories)
                     .Where(f => supportedExtensions.Contains(Path.GetExtension(f), StringComparer.OrdinalIgnoreCase)))
        {
            yield return new object[] { file };
        }
    }

    /// <summary>
    /// Performs a read → write → read round‑trip on the supplied fixture file and
    /// asserts structural equality between the original and the round‑tripped
    /// model.
    /// </summary>
    /// <param name="fixturePath">The absolute path of the fixture file.</param>
    [Theory]
    [MemberData(nameof(FixtureFiles))]
    public void RoundTrip_Fidelity(string fixturePath)
    {
        ArgumentNullException.ThrowIfNull(fixturePath);
        var kind = OfficeDocument.DetectKind(fixturePath);

        // Load the original model.
        var original = kind switch
        {
            DocumentKind.Workbook => (object)OfficeDocument.OpenWorkbook(fixturePath),
            DocumentKind.Document => (object)OfficeDocument.OpenDocument(fixturePath),
            DocumentKind.Presentation => (object)OfficeDocument.OpenPresentation(fixturePath),
            _ => throw new NotSupportedException($"Unsupported document kind: {kind}")
        };

        // Write the model to a temporary file.
        var tempPath = Path.Combine(_tempDir, Path.GetFileName(fixturePath));
        switch (kind)
        {
            case DocumentKind.Workbook:
                OfficeDocument.SaveWorkbook((WorkbookModel)original, tempPath);
                var roundtripWb = OfficeDocument.OpenWorkbook(tempPath);
                AssertWorkbookModelsEqual((WorkbookModel)original, roundtripWb);
                break;

            case DocumentKind.Document:
                OfficeDocument.SaveDocument((DocumentModel)original, tempPath);
                var roundtripDoc = OfficeDocument.OpenDocument(tempPath);
                AssertDocumentModelsEqual((DocumentModel)original, roundtripDoc);
                break;

            case DocumentKind.Presentation:
                OfficeDocument.SavePresentation((PresentationModel)original, tempPath);
                var roundtripPres = OfficeDocument.OpenPresentation(tempPath);
                AssertPresentationModelsEqual((PresentationModel)original, roundtripPres);
                break;
        }
    }

    /// <summary>
    /// Asserts that two <see cref="WorkbookModel"/> instances are structurally
    /// equivalent. The comparison checks sheet names, cell counts and the
    /// values of each cell.
    /// </summary>
    /// <param name="expected">The original workbook model.</param>
    /// <param name="actual">The workbook model obtained after the round‑trip.</param>
    private static void AssertWorkbookModelsEqual(WorkbookModel expected, WorkbookModel actual)
    {
        ArgumentNullException.ThrowIfNull(expected);
        ArgumentNullException.ThrowIfNull(actual);

        Assert.Equal(expected.Sheets.Count, actual.Sheets.Count);

        var expectedSheets = expected.Sheets.OrderBy(s => s.Name).ToArray();
        var actualSheets = actual.Sheets.OrderBy(s => s.Name).ToArray();

        for (int i = 0; i < expectedSheets.Length; i++)
        {
            var expSheet = expectedSheets[i];
            var actSheet = actualSheets[i];

            Assert.Equal(expSheet.Name, actSheet.Name);
            Assert.Equal(expSheet.Cells.Count, actSheet.Cells.Count);

            var expCells = expSheet.OrderedCells().ToArray();
            var actCells = actSheet.OrderedCells().ToArray();

            for (int j = 0; j < expCells.Length; j++)
            {
                var (expRef, expVal) = expCells[j];
                var (actRef, actVal) = actCells[j];

                Assert.Equal(expRef, actRef);
                Assert.Equal(expVal.Kind, actVal.Kind);
                Assert.Equal(expVal.Text, actVal.Text);
                Assert.Equal(expVal.Number, actVal.Number);
                Assert.Equal(expVal.Boolean, actVal.Boolean);
                Assert.Equal(expVal.DateTime, actVal.DateTime);
                Assert.Equal(expVal.Formula, actVal.Formula);
            }
        }
    }

    /// <summary>
    /// Asserts that two <see cref="DocumentModel"/> instances are structurally
    /// equivalent. The comparison checks paragraph count, paragraph kinds,
    /// heading levels, and the runs (text and style) inside each paragraph.
    /// </summary>
    /// <param name="expected">The original document model.</param>
    /// <param name="actual">The document model obtained after the round‑trip.</param>
    private static void AssertDocumentModelsEqual(DocumentModel expected, DocumentModel actual)
    {
        ArgumentNullException.ThrowIfNull(expected);
        ArgumentNullException.ThrowIfNull(actual);

        Assert.Equal(expected.Paragraphs.Count, actual.Paragraphs.Count);

        for (int i = 0; i < expected.Paragraphs.Count; i++)
        {
            var expPara = expected.Paragraphs[i];
            var actPara = actual.Paragraphs[i];

            Assert.Equal(expPara.Kind, actPara.Kind);
            Assert.Equal(expPara.HeadingLevel, actPara.HeadingLevel);
            Assert.Equal(expPara.Text, actPara.Text);

            Assert.Equal(expPara.Runs.Count, actPara.Runs.Count);
            for (int r = 0; r < expPara.Runs.Count; r++)
            {
                var expRun = expPara.Runs[r];
                var actRun = actPara.Runs[r];

                Assert.Equal(expRun.Text, actRun.Text);
                Assert.Equal(expRun.Style.Bold, actRun.Style.Bold);
                Assert.Equal(expRun.Style.Italic, actRun.Style.Italic);
                Assert.Equal(expRun.Style.Underline, actRun.Style.Underline);
                Assert.Equal(expRun.Style.StrikeThrough, actRun.Style.StrikeThrough);
                // Add more style fields here if the model expands.
            }
        }
    }

    /// <summary>
    /// Asserts that two <see cref="PresentationModel"/> instances are structurally
    /// equivalent. The comparison uses the plain‑text representation of each slide
    /// because the underlying shape model is complex and not needed for the
    /// fidelity guarantee exercised by the current test suite.
    /// </summary>
    /// <param name="expected">The original presentation model.</param>
    /// <param name="actual">The presentation model obtained after the round‑trip.</param>
    private static void AssertPresentationModelsEqual(PresentationModel expected, PresentationModel actual)
    {
        ArgumentNullException.ThrowIfNull(expected);
        ArgumentNullException.ThrowIfNull(actual);

        Assert.Equal(expected.Slides.Count, actual.Slides.Count);

        for (int i = 0; i < expected.Slides.Count; i++)
        {
            var expSlide = expected.Slides[i];
            var actSlide = actual.Slides[i];

            // Compare slide titles (if any) and the plain‑text content.
            Assert.Equal(expSlide.Title?.Text ?? string.Empty,
                         actSlide.Title?.Text ?? string.Empty);

            Assert.Equal(expSlide.ToPlainText(), actSlide.ToPlainText());
        }
    }
}
