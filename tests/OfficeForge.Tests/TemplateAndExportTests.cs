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
        Assert.Equal("Dear Ada, welcome to Acme!", filler.FillText("Dear {{name}}, welcome to {{company}}!"));
    }

    [Fact]
    public void FillText_ToleratesWhitespaceInsideBraces()
    {
        Assert.Equal("Ada", Filler(("name", "Ada")).FillText("{{ name }}"));
    }

    [Fact]
    public void FillText_LeavesUnknownPlaceholdersIntact()
    {
        Assert.Equal("Hi {{missing}}", Filler(("name", "Ada")).FillText("Hi {{missing}}"));
    }

    [Fact]
    public void Fill_Document_ReplacesAcrossRunsKeepingFirstStyle()
    {
        var doc = new DocumentModel();
        var paragraph = new ParagraphModel();
        paragraph.Runs.Add(new RunModel("Total: {{tot", new RunStyle(Bold: true)));
        paragraph.Runs.Add(new RunModel("al}} EUR", RunStyle.Default));
        doc.Paragraphs.Add(paragraph);
        Filler(("total", "99")).Fill(doc);
        var run = Assert.Single(paragraph.Runs);
        Assert.Equal("Total: 99 EUR", run.Text);
        Assert.True(run.Style.Bold);
    }

    [Fact]
    public void Fill_Document_LeavesUntouchedParagraphsAlone()
    {
        var doc = new DocumentModel();
        doc.AddParagraph("static text", new RunStyle(Italic: true));
        Filler(("x", "y")).Fill(doc);
        Assert.Equal("static text", doc.Paragraphs[0].Text);
        Assert.True(doc.Paragraphs[0].Runs[0].Style.Italic);
    }

    [Fact]
    public void Fill_Workbook_ReplacesOnlyTextCells()
    {
        var workbook = new WorkbookModel();
        var sheet = workbook.AddSheet("S");
        sheet["A1"] = CellValue.FromText("{{customer}}");
        sheet["B1"] = CellValue.FromNumber(42);
        sheet["C1"] = CellValue.FromFormula("A1&\"{{customer}}\"");
        Filler(("customer", "Acme")).Fill(workbook);
        Assert.Equal("Acme", sheet["A1"].Text);
        Assert.Equal(42, sheet["B1"].Number);
        Assert.Equal("A1&\"{{customer}}\"", sheet["C1"].Formula);
    }

    [Fact]
    public void ParsePairs_SplitsOnFirstEquals()
    {
        var values = TemplateFiller.ParsePairs(["key=a=b", "other=x"]);
        Assert.Equal("a=b", values["key"]);
        Assert.Equal("x", values["other"]);
    }

    [Fact]
    public void ParsePairs_RejectsMalformedInput()
    {
        Assert.Throws<FormatException>(() => TemplateFiller.ParsePairs(["novalue"]));
        Assert.Throws<FormatException>(() => TemplateFiller.ParsePairs(["=leading"]));
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
        sheet["B2"] = CellValue.FromNumber(150);
        return workbook;
    }

    [Fact]
    public void Workbook_ToMarkdown_RendersHeaderAndRows()
    {
        var md = WorkbookExporter.ToMarkdown(SampleWorkbook());
        Assert.Contains("## Sales", md);
        Assert.Contains("| Region | Total |", md);
        Assert.Contains("| --- | --- |", md);
        Assert.Contains("| North | 150 |", md);
    }

    [Fact]
    public void Workbook_ToJson_UsesA1Keys()
    {
        var json = WorkbookExporter.ToJson(SampleWorkbook());
        Assert.Contains("\"Sales\"", json);
        Assert.Contains("\"A2\": \"North\"", json);
        Assert.Contains("\"B2\": \"150\"", json);
    }

    [Fact]
    public void Workbook_ToPlainText_UsesTabSeparatedRows()
    {
        var text = WorkbookExporter.ToPlainText(SampleWorkbook());
        Assert.Contains("Region\tTotal", text);
        Assert.Contains("North\t150", text);
    }

    [Fact]
    public void Document_ToMarkdown_RendersHeadingsListsAndEmphasis()
    {
        var doc = new DocumentModel();
        var h = doc.AddParagraph("Title");
        h.Kind = ParagraphKind.Heading;
        h.HeadingLevel = 2;
        var li = doc.AddParagraph("point", new RunStyle(Italic: true));
        li.Kind = ParagraphKind.ListItem;
        doc.AddParagraph("strong", new RunStyle(Bold: true));
        var md = DocumentExporter.ToMarkdown(doc);
        Assert.Contains("## Title", md);
        Assert.Contains("- *point*", md);
        Assert.Contains("**strong**", md);
    }

    [Fact]
    public void Document_ToJson_IncludesKindAndText()
    {
        var doc = new DocumentModel();
        doc.AddParagraph("hello");
        var json = DocumentExporter.ToJson(doc);
        Assert.Contains("\"kind\": \"Body\"", json);
        Assert.Contains("\"text\": \"hello\"", json);
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
        Assert.Equal(a1, CellRef.Parse(a1).ToA1());
    }

    [Theory]
    [InlineData("1A")]
    [InlineData("A0")]
    [InlineData("A")]
    [InlineData("42")]
    public void TryParse_RejectsInvalidReferences(string input)
    {
        Assert.False(CellRef.TryParse(input, out _));
    }
}
