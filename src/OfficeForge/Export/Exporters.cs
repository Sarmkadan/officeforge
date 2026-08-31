using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.Json;
using OfficeForge.Models;
using OfficeForge.Spreadsheets;
using OfficeForge.Words;
using OfficeForge.Slides;

namespace OfficeForge.Export;

/// <summary>
/// Supported export formats.
/// </summary>
public enum ExportFormat
{
    Text,
    Markdown,
    Json
}

/// <summary>
/// Exports a <see cref="WorkbookModel"/> to various textual representations.
/// </summary>
[Obsolete("Use the generic IDocumentExporter<TModel> implementation instead. This class is kept for compatibility.")]
public static class WorkbookExporter
{
    /// <summary>
    /// Exports the supplied <paramref name="workbook"/> using the specified <paramref name="format"/>.
    /// </summary>
    /// <param name="workbook">The workbook to export.</param>
    /// <param name="format">The desired export format.</param>
    /// <returns>The exported string.</returns>
    /// <exception cref="ArgumentNullException">Thrown when <paramref name="workbook"/> is <c>null</c>.</exception>
    public static string Export(WorkbookModel workbook, ExportFormat format)
    {
        ArgumentNullException.ThrowIfNull(workbook);
        ExporterValidation.ValidateModel(workbook, nameof(workbook));
        return Exporter.Export(workbook, format, ToPlainText, ToMarkdown, ToJson);
    }

    /// <summary>
    /// Reads a workbook from <paramref name="path"/> and exports it using <paramref name="format"/>.
    /// </summary>
    /// <param name="path">The path to the workbook file.</param>
    /// <param name="format">The desired export format.</param>
    /// <returns>The exported string.</returns>
    /// <exception cref="ArgumentNullException">Thrown when <paramref name="path"/> is <c>null</c>.</exception>
    public static string Export(string path, ExportFormat format)
    {
        return Exporter.Export(path, format, new XlsxReader(), Export);
    }

    /// <summary>
    /// Reads a workbook from <paramref name="stream"/> and exports it using <paramref name="format"/>.
    /// </summary>
    /// <param name="stream">The stream containing the workbook data.</param>
    /// <param name="format">The desired export format.</param>
    /// <returns>The exported string.</returns>
    /// <exception cref="ArgumentNullException">Thrown when <paramref name="stream"/> is <c>null</c>.</exception>
    public static string Export(Stream stream, ExportFormat format)
    {
        return Exporter.Export(stream, format, new XlsxReader(), Export);
    }

    /// <summary>
    /// Exports the workbook as plain text (tab‑separated values).
    /// </summary>
    /// <param name="workbook">The workbook to export.</param>
    /// <returns>Plain‑text representation.</returns>
    /// <exception cref="ArgumentNullException">Thrown when <paramref name="workbook"/> is <c>null</c>.</exception>
    public static string ToPlainText(WorkbookModel workbook)
    {
        ArgumentNullException.ThrowIfNull(workbook);
        ExporterValidation.ValidateModel(workbook, nameof(workbook));
        var sb = new StringBuilder();
        foreach (var sheet in workbook.Sheets)
        {
            sb.AppendLine(sheet.Name);
            foreach (var row in Rows(sheet))
                sb.AppendLine(string.Join('\t', row));
        }
        return sb.ToString();
    }

    /// <summary>
    /// Exports the workbook as Markdown tables.
    /// </summary>
    /// <param name="workbook">The workbook to export.</param>
    /// <returns>Markdown representation.</returns>
    /// <exception cref="ArgumentNullException">Thrown when <paramref name="workbook"/> is <c>null</c>.</exception>
    public static string ToMarkdown(WorkbookModel workbook)
    {
        ArgumentNullException.ThrowIfNull(workbook);
        ExporterValidation.ValidateModel(workbook, nameof(workbook));
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

    /// <summary>
    /// Exports the workbook as JSON.
    /// </summary>
    /// <param name="workbook">The workbook to export.</param>
    /// <returns>JSON representation.</returns>
    /// <exception cref="ArgumentNullException">Thrown when <paramref name="workbook"/> is <c>null</c>.</exception>
    public static string ToJson(WorkbookModel workbook)
    {
        ArgumentNullException.ThrowIfNull(workbook);
        ExporterValidation.ValidateModel(workbook, nameof(workbook));
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
                row[cellRef.Column - 1] = value.ToString(); // invariant via CellValue
            yield return row;
        }
    }

    private static IEnumerable<string> Pad(string[] row, int width) =>
        Enumerable.Range(0, width).Select(i => i < row.Length ? row[i].Replace("|", "\\|") : string.Empty);
}

/// <summary>
/// Exports a <see cref="DocumentModel"/> to various textual representations.
/// </summary>
[Obsolete("Use the generic IDocumentExporter<TModel> implementation instead. This class is kept for compatibility.")]
public static class DocumentExporter
{
    /// <summary>
    /// Exports the supplied <paramref name="document"/> using the specified <paramref name="format"/>.
    /// </summary>
    /// <param name="document">The document to export.</param>
    /// <param name="format">The desired export format.</param>
    /// <returns>The exported string.</returns>
    /// <exception cref="ArgumentNullException">Thrown when <paramref name="document"/> is <c>null</c>.</exception>
    public static string Export(DocumentModel document, ExportFormat format)
    {
        ExporterValidation.ValidateModel(document, nameof(document));
        return Exporter.Export(document, format, static model => model.ToPlainText(), ToMarkdown, ToJson);
    }

    /// <summary>
    /// Reads a document from <paramref name="path"/> and exports it using <paramref name="format"/>.
    /// </summary>
    /// <param name="path">The path to the document file.</param>
    /// <param name="format">The desired export format.</param>
    /// <returns>The exported string.</returns>
    /// <exception cref="ArgumentNullException">Thrown when <paramref name="path"/> is <c>null</c>.</exception>
    public static string Export(string path, ExportFormat format)
    {
        return Exporter.Export(path, format, new DocxReader(), Export);
    }

    /// <summary>
    /// Reads a document from <paramref name="stream"/> and exports it using <paramref name="format"/>.
    /// </summary>
    /// <param name="stream">The stream containing the document data.</param>
    /// <param name="format">The desired export format.</param>
    /// <returns>The exported string.</returns>
    /// <exception cref="ArgumentNullException">Thrown when <paramref name="stream"/> is <c>null</c>.</exception>
    public static string Export(Stream stream, ExportFormat format)
    {
        return Exporter.Export(stream, format, new DocxReader(), Export);
    }

    /// <summary>
    /// Exports the document as Markdown.
    /// </summary>
    /// <param name="document">The document to export.</param>
    /// <returns>Markdown representation.</returns>
    /// <exception cref="ArgumentNullException">Thrown when <paramref name="document"/> is <c>null</c>.</exception>
    public static string ToMarkdown(DocumentModel document)
    {
        ExporterValidation.ValidateModel(document, nameof(document));
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

    /// <summary>
    /// Exports the document as JSON.
    /// </summary>
    /// <param name="document">The document to export.</param>
    /// <returns>JSON representation.</returns>
    /// <exception cref="ArgumentNullException">Thrown when <paramref name="document"/> is <c>null</c>.</exception>
    public static string ToJson(DocumentModel document)
    {
        ExporterValidation.ValidateModel(document, nameof(document));
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

/// <summary>
/// Exports a <see cref="PresentationModel"/> to various textual representations.
/// </summary>
[Obsolete("Use the generic IDocumentExporter<TModel> implementation instead. This class is kept for compatibility.")]
public static class PresentationExporter
{
    /// <summary>
    /// Exports the supplied <paramref name="presentation"/> using the specified <paramref name="format"/>.
    /// </summary>
    /// <param name="presentation">The presentation to export.</param>
    /// <param name="format">The desired export format.</param>
    /// <returns>The exported string.</returns>
    /// <exception cref="ArgumentNullException">Thrown when <paramref name="presentation"/> is <c>null</c>.</exception>
    public static string Export(PresentationModel presentation, ExportFormat format)
    {
        ExporterValidation.ValidateModel(presentation, nameof(presentation));
        return Exporter.Export(presentation, format, ToPlainText, ToMarkdown, ToJson);
    }

    /// <summary>
    /// Reads a presentation from <paramref name="path"/> and exports it using <paramref name="format"/>.
    /// </summary>
    /// <param name="path">The path to the presentation file.</param>
    /// <param name="format">The desired export format.</param>
    /// <returns>The exported string.</returns>
    /// <exception cref="ArgumentNullException">Thrown when <paramref name="path"/> is <c>null</c>.</exception>
    public static string Export(string path, ExportFormat format)
    {
        return Exporter.Export(path, format, new PptxReader(), Export);
    }

    /// <summary>
    /// Reads a presentation from <paramref name="stream"/> and exports it using <paramref name="format"/>.
    /// </summary>
    /// <param name="stream">The stream containing the presentation data.</param>
    /// <param name="format">The desired export format.</param>
    /// <returns>The exported string.</returns>
    /// <exception cref="ArgumentNullException">Thrown when <paramref name="stream"/> is <c>null</c>.</exception>
    public static string Export(Stream stream, ExportFormat format)
    {
        return Exporter.Export(stream, format, new PptxReader(), Export);
    }

    /// <summary>
    /// Exports the presentation as plain text (double‑newline separated slides).
    /// </summary>
    /// <param name="presentation">The presentation to export.</param>
    /// <returns>Plain‑text representation.</returns>
    /// <exception cref="ArgumentNullException">Thrown when <paramref name="presentation"/> is <c>null</c>.</exception>
    public static string ToPlainText(PresentationModel presentation)
    {
        ExporterValidation.ValidateModel(presentation, nameof(presentation));
        return string.Join(Environment.NewLine + Environment.NewLine,
            presentation.Slides.Select(s => s.ToPlainText()));
    }

    /// <summary>
    /// Exports the presentation as Markdown.
    /// </summary>
    /// <param name="presentation">The presentation to export.</param>
    /// <returns>Markdown representation.</returns>
    /// <exception cref="ArgumentNullException">Thrown when <paramref name="presentation"/> is <c>null</c>.</exception>
    public static string ToMarkdown(PresentationModel presentation)
    {
        ExporterValidation.ValidateModel(presentation, nameof(presentation));
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

    /// <summary>
    /// Exports the presentation as JSON.
    /// </summary>
    /// <param name="presentation">The presentation to export.</param>
    /// <returns>JSON representation.</returns>
    /// <exception cref="ArgumentNullException">Thrown when <paramref name="presentation"/> is <c>null</c>.</exception>
    public static string ToJson(PresentationModel presentation)
    {
        ExporterValidation.ValidateModel(presentation, nameof(presentation));
        var payload = presentation.Slides.Select(s => new
        {
            title = s.Title,
            shapes = s.Shapes.Select(sh => new { name = sh.Name, lines = sh.Lines })
        });
        return JsonSerializer.Serialize(payload, JsonOptions.Indented);
    }
}

internal static class Exporter
{
    public static string Export<TModel>(
        TModel model,
        ExportFormat format,
        Func<TModel, string> text,
        Func<TModel, string> markdown,
        Func<TModel, string> json) => format switch
        {
            ExportFormat.Markdown => markdown(model),
            ExportFormat.Json => json(model),
            _ => text(model)
        };

    public static string Export<TModel>(
        string path,
        ExportFormat format,
        IDocumentReader<TModel> reader,
        Func<TModel, ExportFormat, string> export)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(path);
        return export(reader.Read(path), format);
    }

    public static string Export<TModel>(
        Stream stream,
        ExportFormat format,
        IDocumentReader<TModel> reader,
        Func<TModel, ExportFormat, string> export)
    {
        ArgumentNullException.ThrowIfNull(stream);
        return export(reader.Read(stream), format);
    }
}

internal static class JsonOptions
{
    public static readonly JsonSerializerOptions Indented = new() { WriteIndented = true };
}

/// <summary>
/// Centralised validation helpers for exporter classes.
/// </summary>
internal static class ExporterValidation
{
    /// <summary>
    /// Validates that the supplied argument is not <c>null</c>.
    /// </summary>
    /// <typeparam name="T">The type of the argument.</typeparam>
    /// <param name="value">The argument value.</param>
    /// <param name="paramName">The name of the argument.</param>
    /// <exception cref="ArgumentNullException">Thrown when <paramref name="value"/> is <c>null</c>.</exception>
    public static void ValidateModel<T>(T value, string paramName) where T : class
    {
        ArgumentNullException.ThrowIfNull(value, paramName);
    }
}
