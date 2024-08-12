# WorkbookModel

The `WorkbookModel` class serves as the primary container for spreadsheet documents within the `officeforge` library. It manages the collection of worksheets and provides an interface for programmatic access to the structural components and cell data contained within a workbook.

## API

### WorkbookModel

*   **`List<SheetModel> Sheets`**
    A list containing all `SheetModel` instances within the workbook. This property allows direct iteration over the worksheets or structural modifications to the workbook's sheet hierarchy.

*   **`SheetModel? FindSheet(string name)`**
    Searches for a worksheet by its name. Returns the `SheetModel` if a match is found; otherwise, returns `null`.

*   **`SheetModel AddSheet(string name)`**
    Creates a new worksheet with the specified name, appends it to the `Sheets` list, and returns the newly created `SheetModel`.

### SheetModel

*   **`string Name`**
    The name of the worksheet.

*   **`Dictionary<CellRef, CellValue> Cells`**
    A dictionary mapping cell references (`CellRef`) to their respective cell values (`CellValue`). This provides $O(1)$ access time for retrieving or updating the value of a specific cell by its reference.

*   **`IEnumerable<KeyValuePair<CellRef, CellValue>> OrderedCells`**
    An enumerable collection providing access to cells in a defined order. This is preferred over `Cells` when the objective is to perform sequential processing, such as serialization or document export.

## Usage

### Creating a Workbook and Adding a Sheet

```csharp
using OfficeForge;

var workbook = new WorkbookModel();
var sheet = workbook.AddSheet("SalesData");

// Populate the sheet
var cellRef = new CellRef("A1");
sheet.Cells[cellRef] = new CellValue("Revenue");
```

### Finding and Processing Data in a Sheet

```csharp
using OfficeForge;

// Assuming 'workbook' is already populated
var sheet = workbook.FindSheet("SalesData");

if (sheet != null)
{
    foreach (var kvp in sheet.OrderedCells)
    {
        Console.WriteLine($"Cell {kvp.Key} has value: {kvp.Value}");
    }
}
```

## Notes

*   **Thread Safety:** Instances of `WorkbookModel` and `SheetModel` are not thread-safe. Concurrent access to the `Sheets` list or the `Cells` dictionary should be synchronized externally if the workbook is being accessed by multiple threads.
*   **Dictionary Lookup:** The `Cells` dictionary uses `CellRef` as a key. Ensure that `CellRef` implements `GetHashCode` and `Equals` correctly to avoid lookup issues or unexpected behavior when inserting or accessing cell data.
*   **Sheet Naming:** While `AddSheet` allows naming, the implementation does not inherently enforce unique names among worksheets. If naming collisions are possible, check for existing sheets using `FindSheet` before adding.
