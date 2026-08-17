// SPDX-License-Identifier: MIT
// Tests for OfficeForge.Templates.TemplateFiller
// Uses the test framework already referenced in the solution (xUnit).

using System;
using System.Collections.Generic;
using OfficeForge.Templates;
using Xunit;

namespace OfficeForge.Tests;

public sealed class TemplateFillerTests
{
    [Fact]
    public void Constructor_NullValues_ThrowsArgumentNullException()
    {
        // Arrange & Act
        var ex = Assert.Throws<ArgumentNullException>(() => new TemplateFiller(null!));

        // Assert
        Assert.Equal("values", ex.ParamName);
    }

    [Fact]
    public void FillText_NullInput_ThrowsArgumentNullException()
    {
        var filler = new TemplateFiller(new Dictionary<string, string>());
        var ex = Assert.Throws<ArgumentNullException>(() => filler.FillText(null!));
        Assert.Equal("text", ex.ParamName);
    }

    [Fact]
    public void FillText_ReplacesPlaceholdersAndEscapesXml()
    {
        // Arrange
        var values = new Dictionary<string, string>(StringComparer.Ordinal)
        {
            ["name"] = "John & Jane",
            ["city"] = "New <York>"
        };
        var filler = new TemplateFiller(values);
        var input = "Hello {{ name }}, welcome to {{ city }}!";

        // Act
        var result = filler.FillText(input);

        // Assert
        // SecurityElement.Escape converts &, <, >, \" and ' to XML entities.
        var expected = "Hello John &amp; Jane, welcome to New &lt;York&gt;!";
        Assert.Equal(expected, result);
    }

    [Fact]
    public void FillText_UnknownPlaceholder_IsLeftUnchanged()
    {
        var values = new Dictionary<string, string>(StringComparer.Ordinal)
        {
            ["known"] = "value"
        };
        var filler = new TemplateFiller(values);
        var input = "This {{ known }} is known, but {{ unknown }} stays.";

        var result = filler.FillText(input);

        var expected = "This value is known, but {{ unknown }} stays.";
        Assert.Equal(expected, result);
    }

    [Fact]
    public void ParsePairs_ValidPairs_ReturnsDictionary()
    {
        // Arrange
        var pairs = new[]
        {
            "key1=value1",
            "key2=123",
            "empty="
        };

        // Act
        var dict = TemplateFiller.ParsePairs(pairs);

        // Assert
        Assert.Equal(3, dict.Count);
        Assert.Equal("value1", dict["key1"]);
        Assert.Equal("123", dict["key2"]);
        Assert.Equal(string.Empty, dict["empty"]);
    }

    [Fact]
    public void ParsePairs_NullInput_ThrowsArgumentNullException()
    {
        var ex = Assert.Throws<ArgumentNullException>(() => TemplateFiller.ParsePairs(null!));
        Assert.Equal("pairs", ex.ParamName);
    }

    [Fact]
    public void ParsePairs_InvalidFormat_ThrowsFormatException()
    {
        var pairs = new[] { "noequalsign" };
        var ex = Assert.Throws<FormatException>(() => TemplateFiller.ParsePairs(pairs));
        Assert.Contains("Expected key=value", ex.Message);
    }
}
