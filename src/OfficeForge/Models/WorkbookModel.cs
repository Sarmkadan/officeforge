namespace OfficeForge.Models;

public sealed class WorkbookModel
{
    public List<SheetModel> Sheets { get; } = [];

    public SheetModel? FindSheet(string name) =>
        Sheets.FirstOrDefault(s => string.Equals(s.Name, name, StringComparison.OrdinalIgnoreCase));

    public SheetModel AddSheet(string name)
    {
        var sheet = new SheetModel(name);
        Sheets.Add(sheet);
        return sheet;
    }
}

public sealed class SheetModel(string name)
{
    public string Name { get; set; } = name ?? throw new ArgumentNullException(nameof(name));
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
}
