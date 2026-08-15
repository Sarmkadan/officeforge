namespace OfficeForge.Tests;

public interface IXlsxRoundTripTests
{
    void Dispose();
    void WriteRead_PreservesTypedCellValues();
    void WriteRead_PreservesMultipleSheets();
    void WriteRead_EmptyWorkbookGetsDefaultSheet();
    void Export_FromXlsxPath_ProducesMarkdownTable();
    void MissingCell_ReadsAsEmpty();
}
