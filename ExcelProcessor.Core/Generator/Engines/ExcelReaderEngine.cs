using DocumentFormat.OpenXml.Packaging;
using ExcelProcessor.Abstractions.Generator.Engines;
using ExcelProcessor.Abstractions.Generator.Sheets.Operations;
using ExcelProcessor.Core.Generator.Sheets.Operations;

namespace ExcelProcessor.Core.Generator.Engines
{
    public class ExcelReaderEngine : ExcelEngineBase, IExcelReaderEngine
    {

        internal ExcelReaderEngine(byte[] data)
        {
            LoadFrom(data);
        }


        public IExcelSheetReader GetSheet(string sheetName)
        {
            WorksheetPart worksheetPart = GetWorksheetPart(sheetName);
            if (worksheetPart == null)
                return null;

            return new ExcelSheetReader(spreadSheetDocument.WorkbookPart, worksheetPart);
        }
    }
}
