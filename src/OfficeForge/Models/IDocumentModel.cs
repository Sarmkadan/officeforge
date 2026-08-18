using System.Collections.Generic;

namespace OfficeForge.Models;

public interface IDocumentModel
{
    List<ParagraphModel> Paragraphs { get; }
    ParagraphModel AddParagraph(string text, RunStyle? style = null);
    string ToPlainText();
}
