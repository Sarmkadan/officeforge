using System.Text.RegularExpressions;
using OfficeForge.Models;

namespace OfficeForge.Templates;

public sealed partial class TemplateFiller(IReadOnlyDictionary<string, string> values)
{
    [GeneratedRegex(@"\{\{\s*([\w.-]+)\s*\}\}")]
    private static partial Regex PlaceholderRegex();

    public string FillText(string text) =>
        PlaceholderRegex().Replace(text, m => values.TryGetValue(m.Groups[1].Value, out var v) ? v : m.Value);

    public void Fill(DocumentModel document)
    {
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
        var values = new Dictionary<string, string>(StringComparer.Ordinal);
        foreach (var pair in pairs)
        {
            var separator = pair.IndexOf('=');
            if (separator <= 0)
                throw new FormatException($"Expected key=value, got '{pair}'.");
            values[pair[..separator]] = pair[(separator + 1)..];
        }
        return values;
    }
}
