using OfficeForge.Models;

namespace OfficeForge.Templates;

/// <summary>
/// Defines the contract for filling templates with values.
/// </summary>
public interface ITemplateFiller
{
    /// <summary>
    /// Replaces placeholders in the given text with the values supplied to the filler.
    /// </summary>
    /// <param name="text">The text containing placeholders.</param>
    /// <returns>The text with placeholders replaced.</returns>
    string FillText(string text);

    /// <summary>
    /// Fills the placeholders in a <see cref="DocumentModel"/> using the supplied values.
    /// </summary>
    /// <param name="document">The document to fill.</param>
    void Fill(DocumentModel document);

    /// <summary>
    /// Fills the placeholders in a <see cref="WorkbookModel"/> using the supplied values.
    /// </summary>
    /// <param name="workbook">The workbook to fill.</param>
    void Fill(WorkbookModel workbook);
}
