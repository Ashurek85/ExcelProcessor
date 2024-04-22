using ExcelProcessor.Abstractions.Pointers;

namespace ExcelProcessor.Abstractions.Generator.Sheets.Operations
{
    public interface IExcelSheetWriter : IExcelSheet
    {
        void InsertValue(string value, string styleName = null);
        void InsertValue(decimal value, string styleName = null);
        void InsertValue(int value, string styleName = null);
        void InsertValue(string value, ICellReference cellRef, string styleName = null);
        void InsertFormula(IFormula formula, string styleName = null);
        void InsertImage(byte[] imgData, string imgDescription, long? customWidth = null, long? customHeight = null);
        void MergeRows(int count, string styleName = null);
        void MergeColumns(int count, string styleName = null);
        void Merge(int columns, int rows, string styleName = null);
        void SetRowHeight(double rowHeight);
        void Save();
    }
}
