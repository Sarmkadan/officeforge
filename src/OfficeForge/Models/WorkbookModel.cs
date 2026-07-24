namespace OfficeForge.Models;

public sealed class WorkbookModel
{
    public List<SheetModel> Sheets { get; } = [];

    public SheetModel? FindSheet(string name) =>
        Sheets.FirstOrDefault(s => string.Equals(s.Name, name, StringComparison.OrdinalIgnoreCase));

    public SheetModel AddSheet(string name)
    {
        ArgumentException.ThrowIfNullOrEmpty(name, nameof(name));

        var existingSheet = FindSheet(name);
        if (existingSheet is not null)
        {
            throw new ArgumentException(
                $"A sheet with name '{name}' already exists in the workbook.",
                nameof(name));
        }

        var sheet = new SheetModel(name);
        Sheets.Add(sheet);
        return sheet;
    }
}