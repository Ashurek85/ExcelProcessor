using ExcelProcessor.Abstractions.Generator.ReaderResults;

namespace ExcelProcessor.Abstractions.Generator.Sheets.Operations
{
    public interface IExcelSheetReader<TEntityReaded> : IExcelSheet
        where TEntityReaded : class, new()
    {
        IExcelReaderResult<TEntityReaded> Results { get; }

        string ReadValue();
        DateTime ReadValueAsDateTime(string customError = null);
        int ReadValueAsInteger(string customError = null);
        bool ReadValueAsYesNo(string customError = null);
        void ProcessAllValuesInParallel(int startsAtRow, int columnCount, Action<string[], uint> processAction);
    }
}
