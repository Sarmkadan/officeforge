using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeForge.Models;

namespace OfficeForge.Words;

public sealed class DocxReader : IDocumentReader<DocumentModel>
{
    public DocumentModel Read(string path)
    {
        using var stream = File.OpenRead(path);
        return Read(stream);
    }

    public DocumentModel Read(Stream stream)
    {
        using var document = WordprocessingDocument.Open(stream, false);
        var body = document.MainDocumentPart?.Document.Body ?? throw new InvalidDataException("Document body missing.");
        var model = new DocumentModel();
        foreach (var paragraph in body.Descendants<Paragraph>())
        {
            var paragraphModel = new ParagraphModel();
            var styleId = paragraph.ParagraphProperties?.ParagraphStyleId?.Val?.Value;
            if (styleId is not null && styleId.StartsWith("Heading", StringComparison.OrdinalIgnoreCase))
            {
                paragraphModel.Kind = ParagraphKind.Heading;
                if (int.TryParse(styleId.AsSpan("Heading".Length), out var level))
                    paragraphModel.HeadingLevel = level;
            }
            foreach (var run in paragraph.Descendants<Run>())
            {
                var text = string.Concat(run.Descendants<Text>().Select(t => t.Text));
                if (text.Length == 0) continue;
                var props = run.RunProperties;
                var style = new Models.RunStyle(
                    Bold: props?.Bold is not null,
                    Italic: props?.Italic is not null,
                    Underline: props?.Underline is not null,
                    FontName: props?.RunFonts?.Ascii?.Value);
                paragraphModel.Runs.Add(new RunModel(text, style));
            }
            model.Paragraphs.Add(paragraphModel);
        }
        return model;
    }
}
