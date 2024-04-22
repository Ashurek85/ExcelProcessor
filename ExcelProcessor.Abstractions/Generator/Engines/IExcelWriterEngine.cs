using ExcelProcessor.Abstractions.Generator.Sheets.Definitions;

namespace ExcelProcessor.Abstractions.Generator.Engines
{
    public interface IExcelWriterEngine : IDisposable
    {
        byte[] Create<TDataContext>(IExcelSheetBuilder<TDataContext>[] sheetBuilders, TDataContext dataContext)
            where TDataContext : class;

        void CreateAndCopy<TDataContext>(IExcelSheetBuilder<TDataContext>[] sheetBuilders, TDataContext dataContext, string outputFile)
           where TDataContext : class;
    }
}