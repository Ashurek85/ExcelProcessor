using ExcelProcessor.Abstractions.Generator.Engines;

namespace ExcelProcessor.Abstractions.Generator
{
    public interface IExcelGenerator
    {
        /// <summary>
        /// Create <see cref="IExcelWriterEngine"/> from template
        /// </summary>
        /// <typeparam name="ExcelStyles">Type of <see cref="IExcelStyles"/> </typeparam>
        /// <param name="templateFilePath">Path to template file</param>
        /// <returns>Instance of <see cref="IExcelWriterEngine"/></returns>
        IExcelWriterEngine FromTemplate<ExcelStyles>(string templateFilePath)
            where ExcelStyles : IExcelStyles, new();

        /// <summary>
        /// Create <see cref="IExcelWriterEngine"/> from empty file
        /// </summary>
        /// <typeparam name="ExcelStyles">Type of <see cref="IExcelStyles"/></typeparam>
        /// <returns>Instance of <see cref="IExcelWriterEngine"/></returns>
        IExcelWriterEngine FromEmptyFile<ExcelStyles>()
            where ExcelStyles : IExcelStyles, new();

        /// <summary>
        /// Create <see cref="IExcelWriterEngine"/> from Excel template as byte array
        /// </summary>
        /// <typeparam name="ExcelStyles">Type of <see cref="IExcelStyles"/></typeparam>
        /// <param name="data">Byte array</param>
        /// <returns>Instance of <see cref="IExcelWriterEngine"/></returns>
        IExcelWriterEngine FromByteArray<ExcelStyles>(byte[] data)
            where ExcelStyles : IExcelStyles, new();

        IExcelReaderEngine FromByteArray(byte[] data);
    }
}
