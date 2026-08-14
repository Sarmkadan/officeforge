using System;

namespace OfficeForge.Export;

/// <summary>
/// Defines a unified export API for a document model.
/// </summary>
/// <typeparam name="TModel">The type of the document model.</typeparam>
public interface IDocumentExporter<TModel>
    where TModel : class
{
    /// <summary>
    /// Exports the supplied <paramref name="model"/> using the specified <paramref name="format"/>.
    /// </summary>
    /// <param name="model">The document model to export.</param>
    /// <param name="format">The desired export format.</param>
    /// <returns>The exported string.</returns>
    /// <exception cref="ArgumentNullException">Thrown when <paramref name="model"/> is <c>null</c>.</exception>
    string Export(TModel model, ExportFormat format);
}
