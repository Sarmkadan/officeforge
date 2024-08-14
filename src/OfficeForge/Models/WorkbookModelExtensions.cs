using System;
using System.Collections.Generic;
using System.Linq;

namespace OfficeForge.Models;

/// <summary>
/// Extension methods for <see cref="WorkbookModel"/> that provide common workbook‑wide operations.
/// </summary>
public static class WorkbookModelExtensions
{
    /// <summary>
    /// Returns a read‑only list containing the names of all sheets in the workbook.
    /// </summary>
    /// <param name="workbook">The workbook to query.</param>
    /// <returns>An <see cref="IReadOnlyList{T}"/> of sheet names.</returns>
    /// <exception cref="ArgumentNullException"><paramref name="workbook"/> is <c>null</c>.</exception>
    public static IReadOnlyList<string> GetSheetNames(this WorkbookModel workbook)
    {
        ArgumentNullException.ThrowIfNull(workbook);
        return workbook.Sheets.Select(s => s.Name).ToArray();
    }

    /// <summary>
    /// Retrieves the value of a cell identified by its <see cref="CellRef"/> on a specific sheet.
    /// </summary>
    /// <param name="workbook">The workbook containing the sheet.</param>
    /// <param name="sheetName">The name of the sheet.</param>
    /// <param name="cellRef">The cell reference.</param>
    /// <returns>The <see cref="CellValue"/> if the cell exists; otherwise <c>null</c>.</returns>
    /// <exception cref="ArgumentNullException"><paramref name="workbook"/> or <paramref name="cellRef"/> is <c>null</c>.</exception>
    /// <exception cref="ArgumentException"><paramref name="sheetName"/> is <c>null</c> or empty, or the sheet does not exist.</exception>
    public static CellValue? GetCellValue(this WorkbookModel workbook, string sheetName, CellRef cellRef)
    {
        ArgumentNullException.ThrowIfNull(workbook);
        ArgumentException.ThrowIfNullOrEmpty(sheetName);
        ArgumentNullException.ThrowIfNull(cellRef);

        var sheet = workbook.FindSheet(sheetName)
            ?? throw new ArgumentException($"Sheet '{sheetName}' not found.", nameof(sheetName));

        return sheet.Cells.TryGetValue(cellRef, out var value) ? value : null;
    }

    /// <summary>
    /// Sets the value of a cell on a given sheet, creating the sheet if it does not already exist.
    /// </summary>
    /// <param name="workbook">The workbook to modify.</param>
    /// <param name="sheetName">The name of the sheet.</param>
    /// <param name="cellRef">The cell reference.</param>
    /// <param name="value">The value to assign.</param>
    /// <exception cref="ArgumentNullException"><paramref name="workbook"/>, <paramref name="cellRef"/> or <paramref name="value"/> is <c>null</c>.</exception>
    /// <exception cref="ArgumentException"><paramref name="sheetName"/> is <c>null</c> or empty.</exception>
    public static void SetCellValue(this WorkbookModel workbook, string sheetName, CellRef cellRef, CellValue value)
    {
        ArgumentNullException.ThrowIfNull(workbook);
        ArgumentException.ThrowIfNullOrEmpty(sheetName);
        ArgumentNullException.ThrowIfNull(cellRef);
        ArgumentNullException.ThrowIfNull(value);

        // Find the sheet or create a new one if it does not exist.
        var sheet = workbook.FindSheet(sheetName) ?? workbook.AddSheet(sheetName);
        sheet.Cells[cellRef] = value;
    }

    /// <summary>
    /// Finds all cells across all sheets whose values satisfy a supplied predicate.
    /// </summary>
    /// <param name="workbook">The workbook to search.</param>
    /// <param name="predicate">A function that determines whether a <see cref="CellValue"/> matches.</param>
    /// <returns>
    /// An <see cref="IEnumerable{T}"/> of tuples containing the sheet, cell reference, and matching value.
    /// </returns>
    /// <exception cref="ArgumentNullException"><paramref name="workbook"/> or <paramref name="predicate"/> is <c>null</c>.</exception>
    public static IEnumerable<(SheetModel Sheet, CellRef CellRef, CellValue Value)> FindCells(
        this WorkbookModel workbook,
        Func<CellValue, bool> predicate)
    {
        ArgumentNullException.ThrowIfNull(workbook);
        ArgumentNullException.ThrowIfNull(predicate);

        return workbook.Sheets
            .SelectMany(sheet => sheet.Cells
                .Where(kv => predicate(kv.Value))
                .Select(kv => (Sheet: sheet, CellRef: kv.Key, Value: kv.Value)));
    }
}
