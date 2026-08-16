using System;
using System.Collections.Generic;
using OfficeForge.Templates;
using Xunit;

namespace OfficeForge.Tests;

/// <summary>
/// Provides extension methods for <see cref="TemplateFillerTests"/>.
/// </summary>
public static class TemplateFillerTestsExtensions
{
    /// <summary>
    /// Asserts that filling the provided template using the given filler results in the expected output.
    /// </summary>
    /// <param name="tests">The <see cref="TemplateFillerTests"/> instance.</param>
    /// <param name="filler">The <see cref="TemplateFiller"/> to use.</param>
    /// <param name="template">The template string.</param>
    /// <param name="expected">The expected output string.</param>
    /// <exception cref="ArgumentNullException">Thrown when <paramref name="filler"/> is null.</exception>
    /// <exception cref="ArgumentException">Thrown when <paramref name="template"/> or <paramref name="expected"/> is null or empty.</exception>
    public static void AssertFillText(this TemplateFillerTests tests, TemplateFiller filler, string template, string expected)
    {
        ArgumentNullException.ThrowIfNull(filler);
        ArgumentException.ThrowIfNullOrEmpty(template);
        ArgumentException.ThrowIfNullOrEmpty(expected);
        Assert.Equal(expected, filler.FillText(template));
    }

    /// <summary>
    /// Creates a <see cref="TemplateFiller"/> from the provided key-value pairs.
    /// </summary>
    /// <param name="tests">The <see cref="TemplateFillerTests"/> instance.</param>
    /// <param name="pairs">The collection of key-value pairs.</param>
    /// <returns>A new <see cref="TemplateFiller"/>.</returns>
    /// <exception cref="ArgumentNullException">Thrown when <paramref name="pairs"/> is null.</exception>
    public static TemplateFiller CreateFiller(this TemplateFillerTests tests, IEnumerable<KeyValuePair<string, string>> pairs)
    {
        ArgumentNullException.ThrowIfNull(pairs);
        return new TemplateFiller(new Dictionary<string, string>(pairs));
    }
}
