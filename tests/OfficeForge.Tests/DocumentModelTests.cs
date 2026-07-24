using System;
using System.Collections.Generic;
using System.Linq;
using OfficeForge.Models;
using Xunit;

namespace OfficeForge.Tests;

public class DocumentModelTests
{
    [Fact]
    public void AddParagraph_WithText_AddsParagraphAndRun_WithDefaultStyle()
    {
        // Arrange
        var doc = new DocumentModel();
        var text = "Hello world";

        // Act
        var paragraph = doc.AddParagraph(text);

        // Assert
        Assert.Single(doc.Paragraphs);
        Assert.Same(paragraph, doc.Paragraphs[0]);

        // Paragraph should contain a single run with the supplied text
        Assert.Single(paragraph.Runs);
        var run = paragraph.Runs[0];
        Assert.Equal(text, run.Text);
        Assert.Equal(RunStyle.Default, run.Style);
    }

    [Fact]
    public void AddParagraph_WithCustomStyle_StoresStyleCorrectly()
    {
        // Arrange
        var doc = new DocumentModel();
        var style = new RunStyle(Bold: true, Italic: true, Underline: false, FontName: "Arial", FontSize: 12);

        // Act
        var paragraph = doc.AddParagraph("Styled text", style);

        // Assert
        var run = Assert.Single(paragraph.Runs);
        Assert.Equal(style, run.Style);
        Assert.True(run.Style.Bold);
        Assert.True(run.Style.Italic);
        Assert.Equal("Arial", run.Style.FontName);
        Assert.Equal(12, run.Style.FontSize);
    }

    [Fact]
    public void AddParagraph_NullText_ThrowsArgumentNullException()
    {
        var doc = new DocumentModel();
        Assert.Throws<ArgumentNullException>(() => doc.AddParagraph(null!));
    }

    [Fact]
    public void ToPlainText_MultipleParagraphs_ReturnsJoinedLines()
    {
        // Arrange
        var doc = new DocumentModel();
        doc.AddParagraph("First line");
        doc.AddParagraph("Second line");
        doc.AddParagraph("Third line");

        // Act
        var plain = doc.ToPlainText();

        // Assert
        var expected = string.Join(Environment.NewLine, new[] { "First line", "Second line", "Third line" });
        Assert.Equal(expected, plain);
    }

    [Fact]
    public void ParagraphModel_Text_ConcatenatesMultipleRuns()
    {
        // Arrange
        var doc = new DocumentModel();
        var paragraph = doc.AddParagraph(string.Empty); // start with empty paragraph
        paragraph.Runs.Add(new RunModel("Run1", RunStyle.Default));
        paragraph.Runs.Add(new RunModel("Run2", RunStyle.Default));
        paragraph.Runs.Add(new RunModel("Run3", RunStyle.Default));

        // Act
        var combined = paragraph.Text;

        // Assert
        Assert.Equal("Run1Run2Run3", combined);
    }

    [Fact]
    public void ParagraphModel_DefaultProperties_AreInitializedCorrectly()
    {
        var doc = new DocumentModel();
        var paragraph = doc.AddParagraph("sample");

        Assert.Equal(ParagraphKind.Body, paragraph.Kind);
        Assert.Equal(0, paragraph.HeadingLevel);
    }

    [Fact]
    public void RunStyle_Default_HasExpectedDefaultValues()
    {
        var def = RunStyle.Default;
        Assert.False(def.Bold);
        Assert.False(def.Italic);
        Assert.False(def.Underline);
        Assert.Null(def.FontName);
        Assert.Null(def.FontSize);
    }
}
