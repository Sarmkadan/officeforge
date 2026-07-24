using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using OfficeForge.Models;

namespace OfficeForge.Templates;

/// <summary>
/// Extension methods that add convenient bulk‑operations and helper functionality to <see cref="TemplateFiller"/>.
/// </summary>
public static class TemplateFillerExtensions
{
    private static readonly Regex PlaceholderRegex = new(@"\{\{\s*([\w.-]+)\s*\}\}", RegexOptions.Compiled);

    /// <summary>
    /// Extracts all distinct placeholder keys from the supplied <paramref name="text"/>.
    /// </summary>
    /// <param name="filler">The <see cref="TemplateFiller"/> instance (unused, required for extension method syntax).</param>
    /// <param name="text">The text to scan for placeholders.</param>
    /// <returns>An <see cref="IReadOnlyCollection{String}"/> containing the unique placeholder names.</returns>
    /// <exception cref="ArgumentNullException">Thrown when <paramref name="filler"/> or <paramref name="text"/> is <c>null</c>.</exception>
    public static IReadOnlyCollection<string> ExtractPlaceholders(this TemplateFiller filler, string text)
    {
        ArgumentNullException.ThrowIfNull(filler);
        ArgumentNullException.ThrowIfNull(text);

        var matches = PlaceholderRegex.Matches(text);
        return matches
            .Select(m => m.Groups[1].Value)
            .Distinct(StringComparer.Ordinal)
            .ToArray(); // array implements IReadOnlyCollection<T>
    }

    /// <summary>
    /// Fills every <see cref="DocumentModel"/> in <paramref name="documents"/> using the current filler values.
    /// </summary>
    /// <param name="filler">The <see cref="TemplateFiller"/> instance.</param>
    /// <param name="documents">A collection of documents to be processed.</param>
    /// <exception cref="ArgumentNullException">Thrown when <paramref name="filler"/> or <paramref name="documents"/> is <c>null</c>.</exception>
    public static void FillAll(this TemplateFiller filler, IEnumerable<DocumentModel> documents)
    {
        ArgumentNullException.ThrowIfNull(filler);
        ArgumentNullException.ThrowIfNull(documents);

        foreach (var doc in documents)
        {
            ArgumentNullException.ThrowIfNull(doc);
            filler.Fill(doc);
        }
    }

    /// <summary>
    /// Fills every <see cref="WorkbookModel"/> in <paramref name="workbooks"/> using the current filler values.
    /// </summary>
    /// <param name="filler">The <see cref="TemplateFiller"/> instance.</param>
    /// <param name="workbooks">A collection of workbooks to be processed.</param>
    /// <exception cref="ArgumentNullException">Thrown when <paramref name="filler"/> or <paramref name="workbooks"/> is <c>null</c>.</exception>
    public static void FillAll(this TemplateFiller filler, IEnumerable<WorkbookModel> workbooks)
    {
        ArgumentNullException.ThrowIfNull(filler);
        ArgumentNullException.ThrowIfNull(workbooks);

        foreach (var wb in workbooks)
        {
            ArgumentNullException.ThrowIfNull(wb);
            filler.Fill(wb);
        }
    }
}
