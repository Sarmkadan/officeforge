using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using OfficeForge.Models;

namespace OfficeForge.Templates;

/// <summary>
/// Extension methods that provide additional bulk‑operations and diagnostics for <see cref="TemplateFiller"/>.
/// </summary>
public static class TemplateFillerExtensions
{
    private static readonly Regex PlaceholderRegex = new(@"\{\{\s*([\w.-]+)\s*\}\}", RegexOptions.Compiled);

    /// <summary>
    /// Fills each string in <paramref name="texts"/> using the current <see cref="TemplateFiller"/> values.
    /// </summary>
    /// <param name="filler">The <see cref="TemplateFiller"/> instance.</param>
    /// <param name="texts">The collection of texts to fill.</param>
    /// <returns>An <see cref="IReadOnlyList{String}"/> containing the filled strings, in the same order as <paramref name="texts"/>.</returns>
    /// <exception cref="ArgumentNullException">Thrown when <paramref name="filler"/> or <paramref name="texts"/> is <c>null</c>.</exception>
    public static IReadOnlyList<string> FillAll(this TemplateFiller filler, IEnumerable<string> texts)
    {
        ArgumentNullException.ThrowIfNull(filler);
        ArgumentNullException.ThrowIfNull(texts);
        return texts.Select(filler.FillText).ToArray();
    }

    /// <summary>
    /// Reads the file at <paramref name="filePath"/>, fills its content using the current <see cref="TemplateFiller"/> values,
    /// and returns the resulting string.
    /// </summary>
    /// <param name="filler">The <see cref="TemplateFiller"/> instance.</param>
    /// <param name="filePath">The path to the template file.</param>
    /// <returns>The filled file content.</returns>
    /// <exception cref="ArgumentNullException">Thrown when <paramref name="filler"/> or <paramref name="filePath"/> is <c>null</c>.</exception>
    /// <exception cref="ArgumentException">Thrown when <paramref name="filePath"/> is an empty string.</exception>
    /// <exception cref="IOException">Thrown when the file cannot be read.</exception>
    public static string FillTemplateFile(this TemplateFiller filler, string filePath)
    {
        ArgumentNullException.ThrowIfNull(filler);
        ArgumentException.ThrowIfNullOrEmpty(filePath);
        var content = File.ReadAllText(filePath);
        return filler.FillText(content);
    }

    /// <summary>
    /// Returns the placeholder keys that remain unreplaced after filling <paramref name="text"/>.
    /// </summary>
    /// <param name="filler">The <see cref="TemplateFiller"/> instance.</param>
    /// <param name="text">The text to analyse.</param>
    /// <returns>A read‑only collection of placeholder names that were not substituted.</returns>
    /// <exception cref="ArgumentNullException">Thrown when <paramref name="filler"/> or <paramref name="text"/> is <c>null</c>.</exception>
    public static IReadOnlyCollection<string> GetUnfilledPlaceholders(this TemplateFiller filler, string text)
    {
        ArgumentNullException.ThrowIfNull(filler);
        ArgumentNullException.ThrowIfNull(text);
        var filled = filler.FillText(text);
        var remaining = PlaceholderRegex
            .Matches(filled)
            .Select(m => m.Groups[1].Value)
            .Distinct(StringComparer.Ordinal)
            .ToArray(); // array implements IReadOnlyCollection<T>
        return remaining;
    }
}
