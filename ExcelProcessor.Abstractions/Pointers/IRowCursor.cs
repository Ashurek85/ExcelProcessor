namespace ExcelProcessor.Abstractions.Pointers
{
    public interface IRowCursor
    {
        ICellReference RowRef { get; }
        ICellReference CellRef { get; }
        void NextColumn();
        void NextRowFromOrigin();
    }
}
