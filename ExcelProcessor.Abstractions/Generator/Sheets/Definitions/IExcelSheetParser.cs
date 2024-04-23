using ExcelProcessor.Abstractions.Generator.ReaderResults;
using ExcelProcessor.Abstractions.Generator.Sheets.Operations;

namespace ExcelProcessor.Abstractions.Generator.Sheets.Definitions
{
    public interface IExcelSheetParser<TEntityReaded>
        where TEntityReaded : class
    {
        /// <summary>
        /// Sheet name in Excel file
        /// </summary>
        string SheetName { get; }

        IExcelReaderResult<TEntityReaded> Parse(IExcelSheetReader sheet);
    }
}
