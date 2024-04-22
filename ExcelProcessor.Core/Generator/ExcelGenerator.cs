using ExcelProcessor.Abstractions;
using ExcelProcessor.Abstractions.Generator;
using ExcelProcessor.Abstractions.Generator.Engines;
using ExcelProcessor.Core.Generator.Engines;

namespace ExcelProcessor.Core.Generator
{
    public class ExcelGenerator : IExcelGenerator
    {
        public IExcelWriterEngine FromTemplate<TExcelStyles>(string templateFilePath)
            where TExcelStyles : IExcelStyles, new()
        {
            return new ExcelWriterEngine<TExcelStyles>(templateFilePath);
        }

        public IExcelWriterEngine FromEmptyFile<TExcelStyles>() 
            where TExcelStyles : IExcelStyles, new()
        {
            return new ExcelWriterEngine<TExcelStyles>();
        }

        public IExcelWriterEngine FromByteArray<TExcelStyles>(byte[] data)
            where TExcelStyles : IExcelStyles, new()
        {
            return new ExcelWriterEngine<TExcelStyles>(data);
        }

        public IExcelReaderEngine FromByteArray(byte[] data)
        {
            return new ExcelReaderEngine(data);
        }
        
    }
}
