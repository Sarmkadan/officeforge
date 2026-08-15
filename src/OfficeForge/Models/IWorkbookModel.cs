namespace OfficeForge.Models
{
    public interface IWorkbookModel
    {
        List<SheetModel> Sheets { get; }
        SheetModel? FindSheet(string name);
        SheetModel AddSheet(string name);
    }
}
