using ExcelProcessor.Abstractions.Generator.Engines;

namespace ExcelProcessor.Abstractions.Generator
{
    public interface IExcelGenerator
    {
        IExcelWriterEngine FromTemplate<ExcelStyles>(string templateFilePath)
            where ExcelStyles : IExcelStyles, new();

        IExcelWriterEngine FromEmptyFile<ExcelStyles>()
            where ExcelStyles : IExcelStyles, new();

        IExcelWriterEngine FromByteArray<ExcelStyles>(byte[] data)
            where ExcelStyles : IExcelStyles, new();

        IExcelReaderEngine FromByteArray(byte[] data);
    }
}
