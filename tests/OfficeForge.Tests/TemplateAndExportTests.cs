using System;
using OfficeForge.Export;
using OfficeForge.Models;
using OfficeForge.Templates;
using Xunit;

namespace OfficeForge.Tests;

public class TemplateFillerTests
{
    static TemplateFiller Filler(params (string Key, string Value)[] pairs) =>
        new(pairs.ToDictionary(p => p.Key, p => p.Value));

    [Fact]
    public void FillText_ReplacesKnownPlaceholders()
    {
        var filler = Filler(("name", "Ada"), ("company", "Acme"));
        Assert.Equal(TemplateFillerTestsConstants.PlaceholderKnown,
            filler.FillText(TemplateFillerTestsConstants.PlaceholderKnown));
    }

    [Fact]
    public void FillText_ToleratesWhitespaceInsideBraces()
    {
        Assert.Equal("Ada",
            Filler(("name", "Ada")).FillText(TemplateFillerTestsConstants.PlaceholderWhitespace));
    }

    [Fact]
    public void FillText_LeavesUnknownPlaceholdersIntact()
    {
        Assert.Equal(TemplateFillerTestsConstants.PlaceholderUnknown,
            Filler(("name", "Ada")).FillText(TemplateFillerTestsConstants.PlaceholderUnknown));
    }

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

    [Fact]
    public void Fill_Document_LeavesUntouchedParagraphsAlone()
    {
        var doc = new DocumentModel();
        doc.AddParagraph(TemplateFillerTestsConstants.StaticText, new RunStyle(Italic: true));
        Filler(("x", "y")).Fill(doc);
        Assert.Equal(TemplateFillerTestsConstants.StaticText, doc.Paragraphs[0].Text);
        Assert.True(doc.Paragraphs[0].Runs[0].Style.Italic);
    }

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

    [Fact]
    public void ParsePairs_RejectsMalformedInput()
    {
        Assert.Throws<FormatException>(() => TemplateFiller.ParsePairs([TemplateFillerTestsConstants.MalformedNoValue]));
        Assert.Throws<FormatException>(() => TemplateFiller.ParsePairs([TemplateFillerTestsConstants.MalformedLeading]));
    }
}

public class ExporterTests
{
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

    [Fact]
    public void Workbook_ToMarkdown_RendersHeaderAndRows()
    {
        var md = WorkbookExporter.ToMarkdown(SampleWorkbook());
        Assert.Contains(TemplateFillerTestsConstants.MarkdownHeader, md);
        Assert.Contains(TemplateFillerTestsConstants.MarkdownHeaderRow, md);
        Assert.Contains(TemplateFillerTestsConstants.MarkdownSeparator, md);
        Assert.Contains(TemplateFillerTestsConstants.MarkdownDataRow, md);
    }

    [Fact]
    public void Workbook_ToJson_UsesA1Keys()
    {
        var json = WorkbookExporter.ToJson(SampleWorkbook());
        Assert.Contains(TemplateFillerTestsConstants.JsonSalesKey, json);
        Assert.Contains(TemplateFillerTestsConstants.JsonA2, json);
        Assert.Contains(TemplateFillerTestsConstants.JsonB2, json);
    }

    [Fact]
    public void Workbook_ToPlainText_UsesTabSeparatedRows()
    {
        var text = WorkbookExporter.ToPlainText(SampleWorkbook());
        Assert.Contains(TemplateFillerTestsConstants.PlainTextHeader, text);
        Assert.Contains(TemplateFillerTestsConstants.PlainTextData, text);
    }

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

public class CellRefTests
{
    [Theory]
    [InlineData("A1", 1, 1)]
    [InlineData("b3", 3, 2)]
    [InlineData("Z10", 10, 26)]
    [InlineData("AA1", 1, 27)]
    [InlineData("AZ5", 5, 52)]
    public void Parse_ResolvesRowAndColumn(string a1, int row, int column)
    {
        ArgumentException.ThrowIfNullOrEmpty(a1);
        var cell = CellRef.Parse(a1);
        Assert.Equal(row, cell.Row);
        Assert.Equal(column, cell.Column);
    }

    [Theory]
    [InlineData("A1")]
    [InlineData("AA99")]
    [InlineData("ZZ1000")]
    public void ToA1_RoundTripsThroughParse(string a1)
    {
        ArgumentException.ThrowIfNullOrEmpty(a1);
        Assert.Equal(a1, CellRef.Parse(a1).ToA1());
    }

    [Theory]
    [InlineData("1A")]
    [InlineData("A0")]
    [InlineData("A")]
    [InlineData("42")]
    public void TryParse_RejectsInvalidReferences(string input)
    {
        ArgumentException.ThrowIfNullOrEmpty(input);
        Assert.False(CellRef.TryParse(input, out _));
    }
}
