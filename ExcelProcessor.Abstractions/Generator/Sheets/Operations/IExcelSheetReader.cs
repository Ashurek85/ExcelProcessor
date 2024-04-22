using ExcelProcessor.Abstractions.Pointers;

namespace ExcelProcessor.Abstractions.Generator.Sheets.Operations
{
    public interface IExcelSheetReader : IExcelSheet
    {        
        string ReadValue(ICellReference cellReference);
        string ReadCursorValue();
        void ProcessAllValuesInParallel(int startsAtRow, int columnCount, Action<string[], uint> processAction);
    }
}
