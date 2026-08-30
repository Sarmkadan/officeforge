using System;
using OfficeForge.Export;
using OfficeForge.Models;
using OfficeForge.Templates;
using Xunit;

namespace OfficeForge.Tests;

/// <summary>
/// Unit tests for <see cref="TemplateFiller"/> covering placeholder replacement in plain text,
/// documents and workbooks, plus parsing of key/value pairs.
/// </summary>
public class TemplateFillerTests
{
    /// <summary>
    /// Creates a <see cref="TemplateFiller"/> from the supplied key/value pairs.
    /// </summary>
    /// <param name="pairs">Placeholder keys paired with the replacement values they map to.</param>
    /// <returns>A filler whose lookup table maps each key to its corresponding value.</returns>
    static TemplateFiller Filler(params (string Key, string Value)[] pairs) =>
        new(pairs.ToDictionary(p => p.Key, p => p.Value));

    /// <summary>
    /// Verifies that <see cref="TemplateFiller.FillText"/> replaces known "{name}" and "{company}"
    /// placeholders with their configured values.
    /// </summary>
    [Fact]
    public void FillText_ReplacesKnownPlaceholders()
    {
        var filler = Filler(("name", "Ada"), ("company", "Acme"));
        Assert.Equal(TemplateFillerTestsConstants.PlaceholderKnown,
            filler.FillText(TemplateFillerTestsConstants.PlaceholderKnown));
    }

    /// <summary>
    /// Verifies that whitespace inside the braces of a placeholder, e.g. "{ name }",
    /// is tolerated and the placeholder still resolves to its configured value.
    /// </summary>
    [Fact]
    public void FillText_ToleratesWhitespaceInsideBraces()
    {
        Assert.Equal("Ada",
            Filler(("name", "Ada")).FillText(TemplateFillerTestsConstants.PlaceholderWhitespace));
    }

    /// <summary>
    /// Verifies that placeholders without a matching entry are left completely intact.
    /// </summary>
    [Fact]
    public void FillText_LeavesUnknownPlaceholdersIntact()
    {
        Assert.Equal(TemplateFillerTestsConstants.PlaceholderUnknown,
            Filler(("name", "Ada")).FillText(TemplateFillerTestsConstants.PlaceholderUnknown));
    }

    /// <summary>
    /// Verifies that filling a document replaces a placeholder split across multiple runs,
    /// merging them into a single run that keeps the style of the first run.
    /// </summary>
    [Fact]
    public void Fill_Document_ReplacesAcrossRunsKeepingFirstStyle()
    {
        var doc = new DocumentModel();
        var paragraph = new ParagraphModel();
        paragraph.Runs.Add(new RunModel(TemplateFillerTestsConstants.RunPart1, new RunStyle(Bold: true)));
        paragraph.Runs.Add(new RunModel(TemplateFillerTestsConstants.RunPart2, RunStyle.Default));
        doc.Paragraphs.Add(paragraph);
        Filler(("total", "99")).Fill(doc);
        var run = Assert.Single(paragraph.Runs);
        Assert.Equal(TemplateFillerTestsConstants.ExpectedRunText, run.Text);
        Assert.True(run.Style.Bold);
    }

    /// <summary>
    /// Verifies that paragraphs containing no placeholders keep their original text and style when a document is filled.
    /// </summary>
    [Fact]
    public void Fill_Document_LeavesUntouchedParagraphsAlone()
    {
        var doc = new DocumentModel();
        doc.AddParagraph(TemplateFillerTestsConstants.StaticText, new RunStyle(Italic: true));
        Filler(("x", "y")).Fill(doc);
        Assert.Equal(TemplateFillerTestsConstants.StaticText, doc.Paragraphs[0].Text);
        Assert.True(doc.Paragraphs[0].Runs[0].Style.Italic);
    }

    /// <summary>
    /// Verifies that filling a workbook replaces placeholders in text cells only,
    /// leaving numeric and formula cells untouched.
    /// </summary>
    [Fact]
    public void Fill_Workbook_ReplacesOnlyTextCells()
    {
        var workbook = new WorkbookModel();
        var sheet = workbook.AddSheet("S");
        sheet["A1"] = CellValue.FromText(TemplateFillerTestsConstants.CustomerPlaceholder);
        sheet["B1"] = CellValue.FromNumber(42);
        sheet["C1"] = CellValue.FromFormula(TemplateFillerTestsConstants.CustomerFormula);
        Filler(("customer", "Acme")).Fill(workbook);
        Assert.Equal("Acme", sheet["A1"].Text);
        Assert.Equal(42, sheet["B1"].Number);
        Assert.Equal(TemplateFillerTestsConstants.CustomerFormula, sheet["C1"].Formula);
    }

    /// <summary>
    /// Verifies that <see cref="TemplateFiller.ParsePairs"/> splits each entry on the first '=',
    /// so values may contain additional '=' characters.
    /// </summary>
    [Fact]
    public void ParsePairs_SplitsOnFirstEquals()
    {
        var values = TemplateFiller.ParsePairs([
            TemplateFillerTestsConstants.InputKeyPair,
            TemplateFillerTestsConstants.InputOtherPair
        ]);
        Assert.Equal(TemplateFillerTestsConstants.ExpectedKeyValue, values["key"]);
        Assert.Equal(TemplateFillerTestsConstants.ExpectedOtherValue, values["other"]);
    }

    /// <summary>
    /// Verifies that entries lacking a value or a key are rejected with a <see cref="FormatException"/>.
    /// </summary>
    [Fact]
    public void ParsePairs_RejectsMalformedInput()
    {
        Assert.Throws<FormatException>(() => TemplateFiller.ParsePairs([TemplateFillerTestsConstants.MalformedNoValue]));
        Assert.Throws<FormatException>(() => TemplateFiller.ParsePairs([TemplateFillerTestsConstants.MalformedLeading]));
    }
}

/// <summary>
/// Unit tests for the workbook and document exporters, verifying Markdown, JSON and plain-text output.
/// </summary>
public class ExporterTests
{
    /// <summary>
    /// Builds a small sales workbook with a header row and a single data row.
    /// </summary>
    /// <returns>A workbook with sheet "Sales" containing Region/Total headers and one North/150 data row.</returns>
    static WorkbookModel SampleWorkbook()
    {
        var workbook = new WorkbookModel();
        var sheet = workbook.AddSheet("Sales");
        sheet["A1"] = CellValue.FromText("Region");
        sheet["B1"] = CellValue.FromText("Total");
        sheet["A2"] = CellValue.FromText("North");
        sheet["B2"] = CellValue.FromNumber(TemplateFillerTestsConstants.Number150);
        return workbook;
    }

    /// <summary>
    /// Verifies that Markdown export renders the header row, the column separator line and the data rows.
    /// </summary>
    [Fact]
    public void Workbook_ToMarkdown_RendersHeaderAndRows()
    {
        var md = WorkbookExporter.ToMarkdown(SampleWorkbook());
        Assert.Contains(TemplateFillerTestsConstants.MarkdownHeader, md);
        Assert.Contains(TemplateFillerTestsConstants.MarkdownHeaderRow, md);
        Assert.Contains(TemplateFillerTestsConstants.MarkdownSeparator, md);
        Assert.Contains(TemplateFillerTestsConstants.MarkdownDataRow, md);
    }

    /// <summary>
    /// Verifies that JSON export nests cells under the sheet name and keys them by A1-style references.
    /// </summary>
    [Fact]
    public void Workbook_ToJson_UsesA1Keys()
    {
        var json = WorkbookExporter.ToJson(SampleWorkbook());
        Assert.Contains(TemplateFillerTestsConstants.JsonSalesKey, json);
        Assert.Contains(TemplateFillerTestsConstants.JsonA2, json);
        Assert.Contains(TemplateFillerTestsConstants.JsonB2, json);
    }

    /// <summary>
    /// Verifies that plain-text export writes tab-separated header and data rows.
    /// </summary>
    [Fact]
    public void Workbook_ToPlainText_UsesTabSeparatedRows()
    {
        var text = WorkbookExporter.ToPlainText(SampleWorkbook());
        Assert.Contains(TemplateFillerTestsConstants.PlainTextHeader, text);
        Assert.Contains(TemplateFillerTestsConstants.PlainTextData, text);
    }

    /// <summary>
    /// Verifies that Markdown export renders headings with hash prefixes, list items with bullets
    /// and bold/italic runs wrapped in emphasis markers.
    /// </summary>
    [Fact]
    public void Document_ToMarkdown_RendersHeadingsListsAndEmphasis()
    {
        var doc = new DocumentModel();
        var h = doc.AddParagraph(TemplateFillerTestsConstants.DocumentTitle);
        h.Kind = ParagraphKind.Heading;
        h.HeadingLevel = TemplateFillerTestsConstants.HeadingLevel;
        var li = doc.AddParagraph(TemplateFillerTestsConstants.DocumentListItem, new RunStyle(Italic: true));
        li.Kind = ParagraphKind.ListItem;
        doc.AddParagraph(TemplateFillerTestsConstants.DocumentStrong, new RunStyle(Bold: true));
        var md = DocumentExporter.ToMarkdown(doc);
        Assert.Contains(TemplateFillerTestsConstants.MarkdownTitle, md);
        Assert.Contains(TemplateFillerTestsConstants.MarkdownListItem, md);
        Assert.Contains(TemplateFillerTestsConstants.MarkdownStrong, md);
    }

    /// <summary>
    /// Verifies that JSON export includes each paragraph's kind and text properties.
    /// </summary>
    [Fact]
    public void Document_ToJson_IncludesKindAndText()
    {
        var doc = new DocumentModel();
        doc.AddParagraph(TemplateFillerTestsConstants.HelloText);
        var json = DocumentExporter.ToJson(doc);
        Assert.Contains(TemplateFillerTestsConstants.JsonKindBody, json);
        Assert.Contains(TemplateFillerTestsConstants.JsonTextHello, json);
    }
}

