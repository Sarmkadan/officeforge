using Xunit;
using OfficeForge.Models;

namespace OfficeForge.Tests;

public class PresentationModelTests
{
    [Fact]
    public void NewPresentationModel_HasEmptySlidesList()
    {
        var presentation = new PresentationModel();
        Assert.Empty(presentation.Slides);
    }

    [Fact]
    public void AddSlide_AddsNewSlideToSlidesList()
    {
        var presentation = new PresentationModel();
        var slide = presentation.AddSlide("Test Slide");
        Assert.Single(presentation.Slides);
        Assert.Same(slide, presentation.Slides[0]);
    }

    [Fact]
    public void AddSlide_NullTitle_DoesNotThrow()
    {
        var presentation = new PresentationModel();
        presentation.AddSlide(null);
        Assert.Single(presentation.Slides);
    }

    [Fact]
    public void ToPlainText_NoSlides_ReturnsEmptyString()
    {
        var presentation = new PresentationModel();
        Assert.Empty(presentation.ToPlainText());
    }

    [Fact]
    public void ToPlainText_SlideWithNoShapes_ReturnsSlideTitle()
    {
        var presentation = new PresentationModel();
        var slide = presentation.AddSlide("Test Slide");
        Assert.Equal("Test Slide", presentation.ToPlainText());
    }

    [Fact]
    public void ToPlainText_SlideWithShapes_ReturnsSlideTitleAndShapes()
    {
        var presentation = new PresentationModel();
        var slide = presentation.AddSlide("Test Slide");
        var shape = new ShapeTextModel { Name = "Test Shape" };
        shape.Lines.Add("Line 1");
        shape.Lines.Add("Line 2");
        slide.Shapes.Add(shape);
        var expected = "Test Slide" + Environment.NewLine + "Line 1" + Environment.NewLine + "Line 2";
        Assert.Equal(expected, presentation.ToPlainText());
    }
}
