using ExcelProcessor.Abstractions.Generator.ReaderResults;
using ExcelProcessor.Abstractions.Generator.Sheets.Definitions;

namespace ExcelProcessor.Abstractions.Generator.Engines
{
    /// <summary>
    /// Excel engine to read data
    /// </summary>
    public interface IExcelReaderEngine : IDisposable
    {
        IExcelReaderResult<TEntityReaded> ReadFile<TEntityReaded>(IExcelSheetParser<TEntityReaded>[] sheetParsers)
            where TEntityReaded : class, new();
    }
}
