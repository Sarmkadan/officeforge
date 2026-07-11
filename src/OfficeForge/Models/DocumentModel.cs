namespace OfficeForge.Models;

public sealed class DocumentModel
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
