using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeForge.Models;

namespace OfficeForge.Words;

public sealed class DocxWriter : IDocumentWriter<DocumentModel>
{
    public void Write(DocumentModel model, string path)
    {
        using var stream = File.Create(path);
        Write(model, stream);
    }

    public void Write(DocumentModel model, Stream stream)
    {
        using var document = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document);
        var mainPart = document.AddMainDocumentPart();
        var body = new Body();
        foreach (var paragraphModel in model.Paragraphs)
        {
            var paragraph = new Paragraph();
            if (paragraphModel.Kind == ParagraphKind.Heading && paragraphModel.HeadingLevel is > 0 and <= 9)
                paragraph.ParagraphProperties = new ParagraphProperties(
                    new ParagraphStyleId { Val = $"Heading{paragraphModel.HeadingLevel}" });
            foreach (var runModel in paragraphModel.Runs)
            {
                var run = new Run();
                var props = new RunProperties();
                if (runModel.Style.Bold) props.Append(new Bold());
                if (runModel.Style.Italic) props.Append(new Italic());
                if (runModel.Style.Underline) props.Append(new Underline { Val = UnderlineValues.Single });
                if (runModel.Style.FontName is { } font) props.Append(new RunFonts { Ascii = font });
                if (props.HasChildren) run.RunProperties = props;
                run.Append(new Text(runModel.Text) { Space = SpaceProcessingModeValues.Preserve });
                paragraph.Append(run);
            }
            body.Append(paragraph);
        }
        mainPart.Document = new Document(body);
        mainPart.Document.Save();
    }
}
