using System.Security;
using System.Text.RegularExpressions;
using OfficeForge.Models;

namespace OfficeForge.Templates;

public sealed partial class TemplateFiller : ITemplateFiller
{
    private readonly IReadOnlyDictionary<string, string> _values;

    public TemplateFiller(IReadOnlyDictionary<string, string> values) =>
        _values = values ?? throw new ArgumentNullException(nameof(values));

    [GeneratedRegex(@"\{\{\s*([\w.-]+)\s*\}\}")]
    private static partial Regex PlaceholderRegex();

    public string FillText(string text)
    {
        ArgumentNullException.ThrowIfNull(text);
        return PlaceholderRegex().Replace(text, m => _values.TryGetValue(m.Groups[1].Value, out var v) ? EscapeXml(v) : m.Value);
    }

    private static string EscapeXml(string value)
    {
        ArgumentNullException.ThrowIfNull(value);
        return SecurityElement.Escape(value);
    }

    public void Fill(DocumentModel document)
    {
        ArgumentNullException.ThrowIfNull(document);
        foreach (var paragraph in document.Paragraphs)
        {
            if (paragraph.Runs.Count == 0) continue;
            var filled = FillText(paragraph.Text);
            if (filled == paragraph.Text) continue;
            var style = paragraph.Runs[0].Style;
            paragraph.Runs.Clear();
            paragraph.Runs.Add(new RunModel(filled, style));
        }
    }

    public void Fill(WorkbookModel workbook)
    {
        ArgumentNullException.ThrowIfNull(workbook);
        foreach (var sheet in workbook.Sheets)
        {
            foreach (var (cellRef, value) in sheet.Cells.ToList())
            {
                if (value.Kind != CellValueKind.Text || value.Text is not { } text) continue;
                var filled = FillText(text);
                if (filled != text) sheet[cellRef] = CellValue.FromText(filled);
            }
        }
    }

    public static IReadOnlyDictionary<string, string> ParsePairs(IEnumerable<string> pairs)
    {
        ArgumentNullException.ThrowIfNull(pairs);
        var values = new Dictionary<string, string>(StringComparer.Ordinal);
        foreach (var pair in pairs)
        {
            ArgumentNullException.ThrowIfNull(pair);
            var separator = pair.IndexOf('=');
            if (separator <= 0)
                throw new FormatException($"Expected key=value, got '{pair}'.");
            values[pair[..separator]] = pair[(separator + 1)..];
        }
        return values;
    }
}
