using System;
using System.Collections.Generic;
using System.IO;
using DocumentFormat.OpenXml.Packaging;
using OfficeForge.Models;
using OfficeForge.Slides;
using Xunit;

namespace OfficeForge.Tests;

public class PptxWriterOptionsTests : IDisposable
{
    readonly string _dir = Directory.CreateTempSubdirectory("officeforge-tests-pptx-options").FullName;

    public void Dispose() => Directory.Delete(_dir, recursive: true);

    public static IEnumerable<object[]> InvalidDimensions =>
    [
        [0],
        [-1],
        [914399],
        [51206401],
        [int.MinValue],
        [int.MaxValue]
    ];

    public static IEnumerable<object[]> InvalidThemeNames =>
    [
        [null!],
        [""],
        ["   "]
    ];

    public static IEnumerable<object[]> InvalidMaxSlides =>
    [
        [0],
        [-1],
        [int.MinValue]
    ];

    [Fact]
    public void Default_HasExpectedValues()
    {
        // Arrange
        var options = PptxWriterOptions.Default;

        // Act & Assert
        Assert.Equal(9144000, options.SlideWidth);
        Assert.Equal(6858000, options.SlideHeight);
        Assert.Equal("Office Theme", options.DefaultTheme);
        Assert.Equal(1000, options.MaxSlides);
        Assert.False(options.IncludeNotes);
    }

    [Fact]
    public void Default_IsValid_WhenValidated()
    {
        // Arrange
        var options = PptxWriterOptions.Default;

        // Act
        var exception = Record.Exception(() => options.Validate());

        // Assert
        Assert.Null(exception);
    }

    [Theory]
    [MemberData(nameof(InvalidDimensions))]
    public void Validate_ThrowsArgumentException_ForSlideWidthOutsideRange(int width)
    {
        // Arrange
        var options = new PptxWriterOptions { SlideWidth = width };

        // Act
        var ex = Assert.Throws<ArgumentException>(() => options.Validate());

        // Assert
        Assert.Equal(nameof(PptxWriterOptions.SlideWidth), ex.ParamName);
    }

    [Theory]
    [MemberData(nameof(InvalidDimensions))]
    public void Validate_ThrowsArgumentException_ForSlideHeightOutsideRange(int height)
    {
        // Arrange
        var options = new PptxWriterOptions { SlideHeight = height };

        // Act
        var ex = Assert.Throws<ArgumentException>(() => options.Validate());

        // Assert
        Assert.Equal(nameof(PptxWriterOptions.SlideHeight), ex.ParamName);
    }

    [Theory]
    [MemberData(nameof(InvalidThemeNames))]
    public void Validate_ThrowsArgumentException_ForNullOrWhitespaceDefaultTheme(string theme)
    {
        // Arrange
        var options = new PptxWriterOptions { DefaultTheme = theme };

        // Act
        var ex = Assert.Throws<ArgumentException>(() => options.Validate());

        // Assert
        Assert.Equal(nameof(PptxWriterOptions.DefaultTheme), ex.ParamName);
    }

    [Theory]
    [MemberData(nameof(InvalidMaxSlides))]
    public void Validate_ThrowsArgumentException_ForNonPositiveMaxSlides(int maxSlides)
    {
        // Arrange
        var options = new PptxWriterOptions { MaxSlides = maxSlides };

        // Act
        var ex = Assert.Throws<ArgumentException>(() => options.Validate());

        // Assert
        Assert.Equal(nameof(PptxWriterOptions.MaxSlides), ex.ParamName);
    }

    [Theory]
    [InlineData(914400, 914400)]
    [InlineData(51206400, 51206400)]
    [InlineData(12192000, 6858000)]
    public void Validate_DoesNotThrow_ForBoundaryDimensions(int width, int height)
    {
        // Arrange
        var options = new PptxWriterOptions { SlideWidth = width, SlideHeight = height };

        // Act
        var exception = Record.Exception(() => options.Validate());

        // Assert
        Assert.Null(exception);
    }

    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void Validate_DoesNotThrow_ForEitherIncludeNotesValue(bool includeNotes)
    {
        // Arrange
        var options = new PptxWriterOptions { IncludeNotes = includeNotes };

        // Act
        var exception = Record.Exception(() => options.Validate());

        // Assert
        Assert.Null(exception);
    }

    [Fact]
    public void PptxWriterConstructor_ThrowsArgumentNullException_ForNullOptions()
    {
        // Arrange, Act & Assert
        Assert.Throws<ArgumentNullException>(() => new PptxWriter(null!));
    }

    [Fact]
    public void PptxWriterConstructor_ThrowsArgumentException_ForInvalidOptions()
    {
        // Arrange
        var options = new PptxWriterOptions { SlideHeight = 1 };

        // Act
        var ex = Assert.Throws<ArgumentException>(() => new PptxWriter(options));

        // Assert
        Assert.Equal(nameof(PptxWriterOptions.SlideHeight), ex.ParamName);
    }

    [Fact]
    public void PptxWriterConstructor_UsesDefaultOptions_WhenParameterless()
    {
        // Arrange & Act
        var writer = new PptxWriter();

        // Assert
        Assert.Same(PptxWriterOptions.Default, writer.Options);
    }

    [Fact]
    public void Write_WritesDefaultSlideSize_ForDefaultOptions()
    {
        // Arrange
        var model = new PresentationModel();
        model.AddSlide("Only slide");
        using var output = new MemoryStream();

        // Act
        new PptxWriter().Write(model, output);

        // Assert
        var (width, height) = ReadSlideSize(output);
        Assert.Equal(9144000, width);
        Assert.Equal(6858000, height);
    }

    [Fact]
    public void Write_WritesConfiguredSlideSize_ToPresentation()
    {
        // Arrange
        var writer = new PptxWriter(new PptxWriterOptions { SlideWidth = 12192000, SlideHeight = 6858000 });
        var model = new PresentationModel();
        model.AddSlide("Widescreen");
        using var output = new MemoryStream();

        // Act
        writer.Write(model, output);

        // Assert
        var (width, height) = ReadSlideSize(output);
        Assert.Equal(12192000, width);
        Assert.Equal(6858000, height);
    }

    [Fact]
    public void Write_ThrowsInvalidOperationException_WhenSlidesExceedMaxSlides()
    {
        // Arrange
        var writer = new PptxWriter(new PptxWriterOptions { MaxSlides = 2 });
        var model = new PresentationModel();
        model.AddSlide("One");
        model.AddSlide("Two");
        model.AddSlide("Three");
        using var output = new MemoryStream();

        // Act
        var ex = Assert.Throws<InvalidOperationException>(() => writer.Write(model, output));

        // Assert
        Assert.Contains(nameof(PptxWriterOptions.MaxSlides), ex.Message);
        Assert.Equal(0, output.Length);
    }

    [Fact]
    public void Write_ToPath_DoesNotCreateFile_WhenSlidesExceedMaxSlides()
    {
        // Arrange
        var path = Path.Combine(_dir, "too-many.pptx");
        var writer = new PptxWriter(new PptxWriterOptions { MaxSlides = 1 });
        var model = new PresentationModel();
        model.AddSlide("One");
        model.AddSlide("Two");

        // Act
        Assert.Throws<InvalidOperationException>(() => writer.Write(model, path));

        // Assert
        Assert.False(File.Exists(path));
    }

    [Fact]
    public void Write_Succeeds_WhenSlideCountEqualsMaxSlides()
    {
        // Arrange
        var writer = new PptxWriter(new PptxWriterOptions { MaxSlides = 2 });
        var model = new PresentationModel();
        model.AddSlide("One");
        model.AddSlide("Two");
        using var output = new MemoryStream();

        // Act
        var exception = Record.Exception(() => writer.Write(model, output));

        // Assert
        Assert.Null(exception);
    }

    private static (int Width, int Height) ReadSlideSize(Stream package)
    {
        package.Position = 0;
        using var document = PresentationDocument.Open(package, isEditable: false);
        var slideSize = document.PresentationPart!.Presentation!.SlideSize!;
        return (slideSize.Cx!.Value, slideSize.Cy!.Value);
    }
}
