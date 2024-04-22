using ExcelProcessor.Abstractions.Generator.Sheets.Operations;

namespace ExcelProcessor.Abstractions.Generator.Sheets.Definitions
{
    public interface IExcelSheetBuilder<TDataContext>
        where TDataContext : class
    {
        string SheetName { get; }
        void Build(IExcelSheetWriter sheet, TDataContext dataContext);
    }
}
