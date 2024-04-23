using ExcelProcessor.Abstractions.Generator.ReaderResults;
using ExcelProcessor.Abstractions.Generator.Sheets.Definitions;
using ExcelProcessor.Abstractions.Generator.Sheets.Operations;
using ExcelProcessor.Examples.Reader.Reader.Entities;

namespace ExcelProcessor.Examples.Reader.Reader
{
    public class ExcelSheetParserExample : IExcelSheetParser<StudentContext>
    {
        public string SheetName => "Students";

        public IExcelReaderResult<StudentContext> Parse(IExcelSheetReader sheet)
        {
            return null;
        }
    }
}
