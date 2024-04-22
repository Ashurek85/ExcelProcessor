using DocumentFormat.OpenXml.Spreadsheet;

namespace ExcelProcessor.Abstractions.Generator.Sheets.Operations
{
    public interface IFormula
    {
        CellFormula Build();
    }
}
