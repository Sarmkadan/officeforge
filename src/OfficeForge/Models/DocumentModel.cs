using System;
using System.Collections.Generic;
using System.Linq;

namespace OfficeForge.Models;

public sealed class DocumentModel : IEquatable<DocumentModel>, IDocumentModel
{
    public List<ParagraphModel> Paragraphs { get; } = [];

    public ParagraphModel AddParagraph(string text, RunStyle? style = null)
    {
        ArgumentNullException.ThrowIfNull(text);
        var paragraph = new ParagraphModel();
        paragraph.Runs.Add(new RunModel(text, style ?? RunStyle.Default));
        Paragraphs.Add(paragraph);
        return paragraph;
    }

    public string ToPlainText() =>
        string.Join(Environment.NewLine, Paragraphs.Select(p => p.Text));

    public bool Equals(DocumentModel? other)
    {
        if (ReferenceEquals(this, other))
            return true;
        if (other is null)
            return false;
        if (Paragraphs.Count != other.Paragraphs.Count)
            return false;
        for (int i = 0; i < Paragraphs.Count; i++)
        {
            var p1 = Paragraphs[i];
            var p2 = other.Paragraphs[i];
            if (p1.Kind != p2.Kind || p1.HeadingLevel != p2.HeadingLevel)
                return false;
        }
        return true;
    }

    public override bool Equals(object? obj) => Equals(obj as DocumentModel);

    public override int GetHashCode()
    {
        var hash = new HashCode();
        foreach (var paragraph in Paragraphs)
        {
            hash.Add(paragraph.Kind);
            hash.Add(paragraph.HeadingLevel);
        }
        return hash.ToHashCode();
    }

    public static bool operator ==(DocumentModel? left, DocumentModel? right) => Equals(left, right);

    public static bool operator !=(DocumentModel? left, DocumentModel? right) => !Equals(left, right);
}

public sealed class ParagraphModel
{
    public List<RunModel> Runs { get; } = [];
    public ParagraphKind Kind { get; set; } = ParagraphKind.Body;
    public int HeadingLevel { get; set; }

    public string Text => string.Concat(Runs.Select(r => r.Text));
}

public enum ParagraphKind
{
    Body,
    Heading,
    ListItem
}

public sealed record RunModel(string Text, RunStyle Style);

public sealed record RunStyle(bool Bold = false, bool Italic = false, bool Underline = false, string? FontName = null, double? FontSize = null)
{
    public static RunStyle Default { get; } = new();
}
