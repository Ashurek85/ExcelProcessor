using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using ExcelProcessor.Abstractions;
using ExcelProcessor.Abstractions.Generator.Sheets.Operations;
using ExcelProcessor.Abstractions.Pointers;
using ExcelProcessor.Core.Exceptions;
using ExcelProcessor.Core.Pointers;

namespace ExcelProcessor.Core.Generator.Sheets.Operations
{
    public class ExcelSheetReader : ExcelSheetBase, IExcelSheetReader
    {
        private readonly WorkbookPart workbookPart;

        public ExcelSheetReader(WorkbookPart workbookPart, WorksheetPart worksheetPart, IExcelStyles styles = null) 
            : base(worksheetPart, styles)
        {
            this.workbookPart = workbookPart ?? throw new ArgumentNullException(nameof(workbookPart));
        }

        public string ReadCursorValue()
        {
            return ReadValue(cursor.CellRef);
        }

        public void ProcessAllValuesInParallel(int startsAtRow, int columnCount, Action<string[], uint> processAction)
        {
            SheetData sheetData = worksheetPart.Worksheet.GetFirstChild<SheetData>();
            var sharedStringTable = workbookPart.GetPartsOfType<SharedStringTablePart>().First().SharedStringTable;

            var rows = sheetData.Elements<Row>().AsEnumerable().Where(r => r.RowIndex >= startsAtRow);
            Parallel.ForEach(rows, row =>
            {
                if (row.RowIndex != null)
                {
                    bool hasData = false;
                    string[] rowData = new string[columnCount];
                    var cells = row.Elements<Cell>();

                    ICellReference cellRef = new CellReference((int)row.RowIndex.Value, "A");
                    int columnIndexToRead = 0;
                    for (int i = 0; i < columnCount; i++)
                    {
                        Cell currentCell = columnIndexToRead < cells.Count() ? cells.ElementAt(columnIndexToRead) : null;
                        if (currentCell != null)
                        {
                            if (currentCell.CellReference.Value == cellRef.ToString())
                            {
                                string cellValue;
                                if (currentCell.DataType != null && currentCell.DataType == CellValues.SharedString)
                                {
                                    int sharedId = int.Parse(currentCell.InnerText);
                                    cellValue = sharedStringTable.ElementAt(sharedId).InnerText;
                                }
                                else
                                    cellValue = currentCell.InnerText;

                                rowData[i] = cellValue;
                                if (!hasData && !string.IsNullOrEmpty(cellValue))
                                    hasData = true;

                                columnIndexToRead++;
                            }
                        }
                        else
                            rowData[i] = null;

                        cellRef = cellRef.NextColumn();
                    }
                    if (hasData)
                        processAction(rowData, row.RowIndex);
                }
            });
        }

        private string ReadValueInternal(ICellReference cellRef)
        {
            Worksheet worksheet = worksheetPart.Worksheet;
            SheetData sheetData = worksheet.GetFirstChild<SheetData>();
            string cellReference = cellRef.ToString();

            // If the worksheet does not contain a row with the specified row index: exception
            Row row;
            if (sheetData.Elements<Row>().Any(r => r.RowIndex == cellRef.Row))
            {
                row = sheetData.Elements<Row>().First(r => r.RowIndex == cellRef.Row);
            }
            else
                throw new RowNotExistsException(cellRef.Row);

            // If there is not a cell with the specified column name, insert one.  
            if (row.Elements<Cell>().Any(c => c.CellReference.Value == cellReference))
            {

                Cell cell = row.Elements<Cell>().First(c => c.CellReference.Value == cellReference);
                if (cell.DataType != null && cell.DataType == CellValues.SharedString)
                {
                    int id;
                    if (int.TryParse(cell.InnerText, out id))
                    {
                        var sharedStringTable = workbookPart.GetPartsOfType<SharedStringTablePart>().FirstOrDefault();
                        if (sharedStringTable != null)
                        {
                            return sharedStringTable.SharedStringTable.ElementAt(id).InnerText;
                        }
                    }
                }
                return cell.CellValue != null ?
                                cell.CellValue.InnerText :
                                null;
            }
            else
                return null;
        }

        public string ReadValue(ICellReference cellReference)
        {
            string value = ReadValueInternal(cellReference);
            return !string.IsNullOrEmpty(value) ? value.Trim() : value;
        }
    }
}
