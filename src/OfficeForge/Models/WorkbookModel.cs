namespace OfficeForge.Models;

public sealed class WorkbookModel : IWorkbookModel
{
    public List<SheetModel> Sheets { get; } = [];

    public SheetModel? FindSheet(string name)
    {
        ArgumentNullException.ThrowIfNull(name, nameof(name));

        return Sheets.FirstOrDefault(s => string.Equals(s.Name, name, StringComparison.OrdinalIgnoreCase));
    }

    public SheetModel AddSheet(string name)
    {
        ArgumentNullException.ThrowIfNull(name, nameof(name));

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
