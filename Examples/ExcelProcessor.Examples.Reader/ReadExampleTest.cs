using ExcelProcessor.Abstractions.Generator;
using ExcelProcessor.Abstractions.Generator.Engines;
using ExcelProcessor.Abstractions.Generator.ReaderResults;
using ExcelProcessor.Core.Generator;
using ExcelProcessor.Examples.Reader.Reader;
using ExcelProcessor.Examples.Reader.Reader.Entities;

namespace ExcelProcessor.Examples.Reader
{
    public class ReadExampleTest
    {
        [Fact]
        public void ReadTest()
        {
            IExcelGenerator excelGenerator = new ExcelGenerator();
            using (IExcelReaderEngine readerEngine = excelGenerator.ReadFromFile("Resources\\ReaderExample.xlsx"))
            {
                IExcelReaderResult<StudentContext> result = readerEngine.ReadFile(new ExcelSheetParserExample());

                Assert.NotNull(result);
            }
        }
    }
}