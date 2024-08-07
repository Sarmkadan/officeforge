using System.Text;
using System.Text.Json;
using OfficeForge.Models;

namespace OfficeForge.Export;

public enum ExportFormat
{
    Text,
    Markdown,
    Json
}

public static class WorkbookExporter
{
    public static string Export(WorkbookModel workbook, ExportFormat format) => format switch
    {
        ExportFormat.Markdown => ToMarkdown(workbook),
        ExportFormat.Json => ToJson(workbook),
        _ => ToPlainText(workbook)
    };

    public static string ToPlainText(WorkbookModel workbook)
    {
        var sb = new StringBuilder();
        foreach (var sheet in workbook.Sheets)
        {
            sb.AppendLine(sheet.Name);
            foreach (var row in Rows(sheet))
                sb.AppendLine(string.Join('\t', row));
        }
        return sb.ToString();
    }

    public static string ToMarkdown(WorkbookModel workbook)
    {
        var sb = new StringBuilder();
        foreach (var sheet in workbook.Sheets)
        {
            sb.AppendLine($"## {sheet.Name}");
            var rows = Rows(sheet).ToList();
            if (rows.Count == 0) continue;
            var width = rows.Max(r => r.Length);
            sb.AppendLine("| " + string.Join(" | ", Pad(rows[0], width)) + " |");
            sb.AppendLine("|" + string.Concat(Enumerable.Repeat(" --- |", width)));
            foreach (var row in rows.Skip(1))
                sb.AppendLine("| " + string.Join(" | ", Pad(row, width)) + " |");
        }
        return sb.ToString();
    }

    public static string ToJson(WorkbookModel workbook)
    {
        var payload = workbook.Sheets.ToDictionary(
            s => s.Name,
            s => s.OrderedCells().ToDictionary(kv => kv.Key.ToA1(), kv => kv.Value.ToString()));
        return JsonSerializer.Serialize(payload, JsonOptions.Indented);
    }

    private static IEnumerable<string[]> Rows(SheetModel sheet)
    {
        var columns = sheet.ColumnCount;
        foreach (var group in sheet.OrderedCells().GroupBy(kv => kv.Key.Row))
        {
            var row = new string[columns];
            Array.Fill(row, string.Empty);
            foreach (var (cellRef, value) in group)
                row[cellRef.Column - 1] = value.ToString();
            yield return row;
        }
    }

    private static IEnumerable<string> Pad(string[] row, int width) =>
        Enumerable.Range(0, width).Select(i => i < row.Length ? row[i].Replace("|", "\\|") : string.Empty);
}

public static class DocumentExporter
{
    public static string Export(DocumentModel document, ExportFormat format) => format switch
    {
        ExportFormat.Markdown => ToMarkdown(document),
        ExportFormat.Json => ToJson(document),
        _ => document.ToPlainText()
    };

    public static string ToMarkdown(DocumentModel document)
    {
        var sb = new StringBuilder();
        foreach (var paragraph in document.Paragraphs)
        {
            var text = string.Concat(paragraph.Runs.Select(FormatRun));
            if (paragraph.Kind == ParagraphKind.Heading)
                text = new string('#', Math.Clamp(paragraph.HeadingLevel, 1, 6)) + " " + text;
            else if (paragraph.Kind == ParagraphKind.ListItem)
                text = "- " + text;
            sb.AppendLine(text);
            sb.AppendLine();
        }
        return sb.ToString();
    }

    public static string ToJson(DocumentModel document)
    {
        var payload = document.Paragraphs.Select(p => new
        {
            kind = p.Kind.ToString(),
            headingLevel = p.HeadingLevel,
            text = p.Text
        });
        return JsonSerializer.Serialize(payload, JsonOptions.Indented);
    }

    private static string FormatRun(RunModel run) => run.Style switch
    {
        { Bold: true, Italic: true } => $"***{run.Text}***",
        { Bold: true } => $"**{run.Text}**",
        { Italic: true } => $"*{run.Text}*",
        _ => run.Text
    };
}

public static class PresentationExporter
{
    public static string Export(PresentationModel presentation, ExportFormat format) => format switch
    {
        ExportFormat.Markdown => ToMarkdown(presentation),
        ExportFormat.Json => ToJson(presentation),
        _ => ToPlainText(presentation)
    };

    public static string ToPlainText(PresentationModel presentation) =>
        string.Join(Environment.NewLine + Environment.NewLine, presentation.Slides.Select(s => s.ToPlainText()));

    public static string ToMarkdown(PresentationModel presentation)
    {
        var sb = new StringBuilder();
        var index = 1;
        foreach (var slide in presentation.Slides)
        {
            sb.AppendLine($"## {slide.Title ?? $"Slide {index}"}");
            foreach (var line in slide.Shapes.SelectMany(s => s.Lines))
                sb.AppendLine("- " + line);
            sb.AppendLine();
            index++;
        }
        return sb.ToString();
    }

    public static string ToJson(PresentationModel presentation)
    {
        var payload = presentation.Slides.Select(s => new
        {
            title = s.Title,
            shapes = s.Shapes.Select(sh => new { name = sh.Name, lines = sh.Lines })
        });
        return JsonSerializer.Serialize(payload, JsonOptions.Indented);
    }
}

internal static class JsonOptions
{
    public static readonly JsonSerializerOptions Indented = new() { WriteIndented = true };
}
