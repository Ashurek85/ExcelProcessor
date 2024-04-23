using ExcelProcessor.Abstractions.Generator.Sheets.Operations;

namespace ExcelProcessor.Abstractions.Generator.Engines
{
    /// <summary>
    /// Excel engine to read data
    /// </summary>
    public interface IExcelReaderEngine : IDisposable
    {
        /// <summary>
        /// Get excel sheet reader instance
        /// </summary>
        /// <param name="sheetName">Sheet name in Excel file</param>
        /// <returns></returns>
        IExcelSheetReader GetSheet(string sheetName);
    }
}
