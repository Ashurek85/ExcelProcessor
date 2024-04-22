using ExcelProcessor.Abstractions.Generator.Sheets.Operations;

namespace ExcelProcessor.Abstractions.Generator.Engines
{
    public interface IExcelReaderEngine : IDisposable
    {
        IExcelSheetReader GetSheet(string sheetName);
    }
}
