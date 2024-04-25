using ExcelProcessor.Abstractions.Generator.ReaderResults;
using ExcelProcessor.Abstractions.Generator.Sheets.Operations;

namespace ExcelProcessor.Abstractions.Generator.Sheets.Definitions
{
    public interface IExcelSheetParser<TEntityReaded>
        where TEntityReaded : class, new()
    {
        /// <summary>
        /// Sheet name in Excel file
        /// </summary>
        string SheetName { get; }

        void Parse(IExcelSheetReader<TEntityReaded> sheet);
    }
}
