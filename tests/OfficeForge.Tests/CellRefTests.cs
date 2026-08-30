using System;
using OfficeForge.Models;
using Xunit;

namespace OfficeForge.Tests;

public class CellRefTests
{
    [Fact]
    public void Parse_SimpleRef_ReturnsCorrectCellRef()
    {
        // Act
        var cell = CellRef.Parse("A1");

        // Assert
        Assert.Equal(1, cell.Row);
        Assert.Equal(1, cell.Column);
    }

    [Fact]
    public void Parse_MultiLetterRef_ReturnsCorrectCellRef()
    {
        // Act
        var cell = CellRef.Parse("AA100");

        // Assert
        Assert.Equal(100, cell.Row);
        Assert.Equal(27, cell.Column); // AA = 26*1 + 1 = 27
    }

    [Fact]
    public void Parse_MaxColumnAndRowRef_ReturnsCorrectCellRef()
    {
        // Act
        var cell = CellRef.Parse("XFD1048576");

        // Assert
        Assert.Equal(CellRef.MaxRows, cell.Row);
        Assert.Equal(CellRef.MaxColumns, cell.Column);
    }

    [Fact]
    public void Parse_LowercaseRef_ReturnsCorrectCellRef()
    {
        // Act
        var cell = CellRef.Parse("b2");

        // Assert
        Assert.Equal(2, cell.Row);
        Assert.Equal(2, cell.Column); // B = 2
    }

    [Fact]
    public void Parse_InvalidFormat_ThrowsFormatException()
    {
        // Arrange
        var invalidInputs = new[] { "1A", "A", "123", "A0", "XFE1" };

        // Act & Assert
        foreach (var input in invalidInputs)
        {
            Assert.Throws<FormatException>(() => CellRef.Parse(input));
        }
    }

    [Fact]
    public void Parse_RowOverMaxRows_ThrowsFormatException()
    {
        // Act & Assert
        Assert.Throws<FormatException>(() => CellRef.Parse("A1048577"));
    }

    [Fact]
    public void Parse_EmptyOrWhitespace_ThrowsArgumentException()
    {
        // Act & Assert
        Assert.Throws<ArgumentException>(() => CellRef.Parse(string.Empty));
        Assert.Throws<ArgumentException>(() => CellRef.Parse("   "));
    }

    [Fact]
    public void TryParse_ValidInput_ReturnsTrueAndCellRef()
    {
        // Act
        var success = CellRef.TryParse("C5", out var cell);

        // Assert
        Assert.True(success);
        Assert.Equal(5, cell.Row);
        Assert.Equal(3, cell.Column);
    }

    [Fact]
    public void TryParse_InvalidInput_ReturnsFalseAndDefault()
    {
        // Act
        var success = CellRef.TryParse("Z", out var cell);

        // Assert
        Assert.False(success);
        Assert.Equal(default, cell);
    }

    [Fact]
    public void ToA1_ToString_RoundTrip_MatchesNormalizedInput()
    {
        // Arrange
        var inputs = new[] { "a1", "B2", "aa100", "XFD1048576" };

        // Act & Assert
        foreach (var input in inputs)
        {
            var cell = CellRef.Parse(input);
            var roundTrip = cell.ToA1();
            var normalized = input.ToUpperInvariant();
            Assert.Equal(normalized, roundTrip);
        }
    }

    [Fact]
    public void ColumnName_ReturnsExpectedStrings()
    {
        // Act & Assert
        Assert.Equal("A", CellRef.ColumnName(1));
        Assert.Equal("Z", CellRef.ColumnName(26));
        Assert.Equal("AA", CellRef.ColumnName(27));
        Assert.Equal("XFD", CellRef.ColumnName(CellRef.MaxColumns));
    }

    [Fact]
    public void ColumnName_Zero_ThrowsArgumentOutOfRangeException()
    {
        // Act & Assert
        Assert.Throws<ArgumentOutOfRangeException>(() => CellRef.ColumnName(0));
    }
}