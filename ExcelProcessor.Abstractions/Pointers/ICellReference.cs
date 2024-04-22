namespace ExcelProcessor.Abstractions.Pointers
{
    public interface ICellReference
    {
        int Row { get; }
        string Column { get; }
        ICellReference NextColumn();
        ICellReference NextRow();
        string ToExcelString();
        string ToDoubleChange(decimal number);
        int GetColumnIndex();
    }
}
