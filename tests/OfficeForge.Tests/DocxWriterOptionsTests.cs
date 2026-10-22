using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Text;
using OfficeForge.Models;
using OfficeForge.Words;
using Xunit;

namespace OfficeForge.Tests;

public class DocxWriterOptionsTests
{
    public static IEnumerable<object[]> InvalidFontNames =>
    [
        [null!],
        [""],
        ["   "]
    ];

    public static IEnumerable<object[]> InvalidFontSizes =>
    [
        [0.0],
        [-1.0],
        [0.5],
        [400.5],
        [401.0],
        [double.NaN],
        [double.PositiveInfinity],
        [double.NegativeInfinity]
    ];

    public static IEnumerable<object[]> InvalidMargins
    {
        get
        {
            foreach (var side in new[] { "Top", "Bottom", "Left", "Right" })
            {
                foreach (var value in new[] { -1.0, -0.01, 1585.0, double.NaN, double.PositiveInfinity })
                {
                    yield return [side, value];
                }
            }
        }
    }

    public static IEnumerable<object[]> InvalidCompressionLevels =>
    [
        [(CompressionLevel)99],
        [(CompressionLevel)(-1)]
    ];

    [Fact]
    public void Default_IsValid_WhenValidated()
    {
        // Arrange
        var options = DocxWriterOptions.Default;

        // Act
        var exception = Record.Exception(() => options.Validate());

        // Assert
        Assert.Null(exception);
    }

    [Theory]
    [MemberData(nameof(InvalidFontNames))]
    public void Validate_ThrowsArgumentException_ForNullOrWhitespaceFontName(string fontName)
    {
        // Arrange
        var options = new DocxWriterOptions { DefaultFontName = fontName };

        // Act
        var ex = Assert.Throws<ArgumentException>(() => options.Validate());

        // Assert
        Assert.Equal(nameof(DocxWriterOptions.DefaultFontName), ex.ParamName);
    }

    [Theory]
    [MemberData(nameof(InvalidFontSizes))]
    public void Validate_ThrowsArgumentException_ForFontSizeOutsideRange(double fontSize)
    {
        // Arrange
        var options = new DocxWriterOptions { DefaultFontSize = fontSize };

        // Act
        var ex = Assert.Throws<ArgumentException>(() => options.Validate());

        // Assert
        Assert.Equal(nameof(DocxWriterOptions.DefaultFontSize), ex.ParamName);
    }

    [Theory]
    [MemberData(nameof(InvalidMargins))]
    public void Validate_ThrowsArgumentException_ForInvalidMargin(string side, double value)
    {
        // Arrange
        var margins = new PageMargins();
        SetMargin(margins, side, value);
        var options = new DocxWriterOptions { PageMargins = margins };

        // Act
        var ex = Assert.Throws<ArgumentException>(() => options.Validate());

        // Assert
        Assert.Equal($"PageMargins.{side}", ex.ParamName);
    }

    [Fact]
    public void Validate_ThrowsArgumentException_ForNullPageMargins()
    {
        // Arrange
        var options = new DocxWriterOptions { PageMargins = null! };

        // Act
        var ex = Assert.Throws<ArgumentException>(() => options.Validate());

        // Assert
        Assert.Equal(nameof(DocxWriterOptions.PageMargins), ex.ParamName);
    }

    [Theory]
    [MemberData(nameof(InvalidCompressionLevels))]
    public void Validate_ThrowsArgumentException_ForUndefinedCompressionLevel(CompressionLevel level)
    {
        // Arrange
        var options = new DocxWriterOptions { CompressionLevel = level };

        // Act
        var ex = Assert.Throws<ArgumentException>(() => options.Validate());

        // Assert
        Assert.Equal(nameof(DocxWriterOptions.CompressionLevel), ex.ParamName);
    }

    [Theory]
    [InlineData(1.0, 0.0)]
    [InlineData(400.0, 1584.0)]
    [InlineData(12.5, 72.0)]
    public void Validate_DoesNotThrow_ForBoundaryValues(double fontSize, double margin)
    {
        // Arrange
        var margins = new PageMargins { Top = margin, Bottom = margin, Left = margin, Right = margin };
        var options = new DocxWriterOptions { DefaultFontSize = fontSize, PageMargins = margins };

        // Act
        var exception = Record.Exception(() => options.Validate());

        // Assert
        Assert.Null(exception);
    }

    [Theory]
    [InlineData(CompressionLevel.Optimal)]
    [InlineData(CompressionLevel.Fastest)]
    [InlineData(CompressionLevel.NoCompression)]
    [InlineData(CompressionLevel.SmallestSize)]
    public void Validate_DoesNotThrow_ForEveryDefinedCompressionLevel(CompressionLevel level)
    {
        // Arrange
        var options = new DocxWriterOptions { CompressionLevel = level };

        // Act
        var exception = Record.Exception(() => options.Validate());

        // Assert
        Assert.Null(exception);
    }

    [Fact]
    public void DocxWriterConstructor_ThrowsArgumentNullException_ForNullOptions()
    {
        // Arrange, Act & Assert
        Assert.Throws<ArgumentNullException>(() => new DocxWriter(null!));
    }

    [Fact]
    public void DocxWriterConstructor_ThrowsArgumentException_ForInvalidOptions()
    {
        // Arrange
        var options = new DocxWriterOptions { DefaultFontSize = 0 };

        // Act
        var ex = Assert.Throws<ArgumentException>(() => new DocxWriter(options));

        // Assert
        Assert.Equal(nameof(DocxWriterOptions.DefaultFontSize), ex.ParamName);
    }

    [Fact]
    public void Write_AppliesDefaultFontAndMargins_ToPackage()
    {
        // Arrange
        var options = new DocxWriterOptions
        {
            DefaultFontName = "Arial",
            DefaultFontSize = 16,
            PageMargins = new PageMargins { Top = 144, Bottom = 144, Left = 36, Right = 36 }
        };
        var model = new DocumentModel();
        model.AddParagraph("Hello");
        using var output = new MemoryStream();

        // Act
        new DocxWriter(options).Write(model, output);

        // Assert
        var styles = ReadEntry(output, "word/styles.xml");
        var document = ReadEntry(output, "word/document.xml");
        Assert.Contains("w:ascii=\"Arial\"", styles);
        Assert.Contains("w:val=\"32\"", styles);
        Assert.Contains("w:top=\"2880\"", document);
        Assert.Contains("w:left=\"720\"", document);
    }

    [Theory]
    [InlineData(CompressionLevel.Optimal)]
    [InlineData(CompressionLevel.Fastest)]
    [InlineData(CompressionLevel.NoCompression)]
    [InlineData(CompressionLevel.SmallestSize)]
    public void Write_RoundTrips_ForEveryCompressionLevel(CompressionLevel level)
    {
        // Arrange
        var model = new DocumentModel();
        model.AddParagraph("Compressed text");
        var writer = new DocxWriter(new DocxWriterOptions { CompressionLevel = level });
        using var output = new MemoryStream();

        // Act
        writer.Write(model, output);
        output.Position = 0;
        var roundTripped = new DocxReader().Read(output);

        // Assert
        Assert.Equal("Compressed text", roundTripped.ToPlainText());
    }

    private static void SetMargin(PageMargins margins, string side, double value)
    {
        switch (side)
        {
            case "Top": margins.Top = value; break;
            case "Bottom": margins.Bottom = value; break;
            case "Left": margins.Left = value; break;
            case "Right": margins.Right = value; break;
            default: throw new ArgumentOutOfRangeException(nameof(side), side, null);
        }
    }

    private static string ReadEntry(Stream package, string entryName)
    {
        package.Position = 0;
        using var archive = new ZipArchive(package, ZipArchiveMode.Read, leaveOpen: true);
        using var reader = new StreamReader(archive.GetEntry(entryName)!.Open(), Encoding.UTF8);
        return reader.ReadToEnd();
    }
}
