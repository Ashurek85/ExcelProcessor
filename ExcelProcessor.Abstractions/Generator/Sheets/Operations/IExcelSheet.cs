using ExcelProcessor.Abstractions.Pointers;

namespace ExcelProcessor.Abstractions.Generator.Sheets.Operations
{
    public interface IExcelSheet
    {
        IRowCursor InitializeCursor(ICellReference cellRef);
    }
}
