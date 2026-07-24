namespace OfficeForge.Models;

/// <summary>
/// Represents a worksheet in an Excel workbook.
/// </summary>
/// <remarks>
/// Sheet names must follow Excel's naming rules:
/// - Maximum 31 characters
/// - Cannot be empty or null
/// - Cannot contain: : \ / ? * [ ]
/// - Cannot contain leading or trailing whitespace
/// </remarks>
public sealed class SheetModel(string name)
{
    private string _name = ValidateSheetName(name);

    public string Name
    {
        get => _name;
        set => _name = ValidateSheetName(value);
    }

    public Dictionary<CellRef, CellValue> Cells { get; } = [];

    public CellValue this[CellRef cell]
    {
        get => Cells.TryGetValue(cell, out var value) ? value : CellValue.Empty;
        set => Cells[cell] = value;
    }

    public CellValue this[string a1]
    {
        get => this[CellRef.Parse(a1)];
        set => this[CellRef.Parse(a1)] = value;
    }

    public int RowCount => Cells.Count == 0 ? 0 : Cells.Keys.Max(c => c.Row);
    public int ColumnCount => Cells.Count == 0 ? 0 : Cells.Keys.Max(c => c.Column);

    public IEnumerable<KeyValuePair<CellRef, CellValue>> OrderedCells() =>
        Cells.OrderBy(kv => kv.Key.Row).ThenBy(kv => kv.Key.Column);

    /// <summary>
    /// Validates a sheet name against Excel's naming rules.
    /// </summary>
    /// <param name="name">The sheet name to validate.</param>
    /// <returns>The validated sheet name.</returns>
    /// <exception cref="ArgumentNullException">Thrown when name is null.</exception>
    /// <exception cref="ArgumentException">Thrown when name violates Excel's sheet naming rules.</exception>
    private static string ValidateSheetName(string name)
    {
        ArgumentNullException.ThrowIfNull(name, nameof(name));

        if (string.IsNullOrWhiteSpace(name))
        {
            throw new ArgumentException(
                "Sheet name cannot be empty or whitespace.",
                nameof(name));
        }

        if (name.Length > 31)
        {
            throw new ArgumentException(
                "Sheet name cannot exceed 31 characters.",
                nameof(name));
        }

        // Check for forbidden characters: : \ / ? * [ ]
        var forbiddenChars = new[] { ':', '\\', '/', '?', '*', '[', ']' };
        if (name.IndexOfAny(forbiddenChars) >= 0)
        {
            throw new ArgumentException(
                "Sheet name cannot contain any of the following characters: : \\ / ? * [ ]",
                nameof(name));
        }

        // Excel also doesn't allow leading or trailing whitespace in sheet names
        if (name.Trim() != name)
        {
            throw new ArgumentException(
                "Sheet name cannot have leading or trailing whitespace.",
                nameof(name));
        }

        return name;
    }
}
