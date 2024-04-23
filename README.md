# ExcelProcessor

Wrapper over OpenXML for ease of use. No paid license required. Offers the functionality to read and generate Excel files

## Positioning
Position and movement through Excel is done through two abstractions:
- ICellReference. Defines a position: row and column
- IRowCursor. Cursor with the ability to change cell reference position. The first position is marked as origin
  - NextColumn(): Move to the next column
  - NextRowFromOrigin(): Allows positioning in the next row and column of origin

## Write operations
Steps:
1. Create a instance of ExcelGenerator. It has methods to generate an instance of IExcelWriterEngine from:
   - Excel Template file
   - Byte array that represents a Excel file
   - Empty file
  
2. Define Excel styles. You need a class that inherits from ExcelStyles. The objective is the definition of styles (InjectStyle method).
   Methods are available to:
   - Inyect fill, border and fonts
   - Generate solid fill with full border, can be customized.

3. Use IExcelWriterEngine to perform operations and get Excel file as byte array (Create method) or save it to disk (CreateAndSave method). Two main parameters:
   - TDataContext. Data context. It will contain the data that you want to write in Excel
   - IExcelSheetBuilder<TDataContext>[]. Each instance of IExcelSheetBuilder must contain:
     - SheetName: the name of the Excel sheet referenced.
     - Implementation of the Build method. The operations to be performed on the Excel sheet will be indicated.

4. Implement IExcelSheetBuilder. Build method usually follow the following steps:
   - Initialize cursor in a cell
   - Perform operations on the cell: InsertValue, InsertFormula, InsertImage, Merge, SetRowHeight, etc
   - User cursor to move to another position: NextColumn or NextRowFromOrigin
   - Continue with perform operations on cell
  
For more information see the example (complete in the source code)
```C#
IExcelGenerator excelGenerator = new ExcelGenerator();
using (IExcelWriterEngine writerEngine = excelGenerator.FromTemplate<ExampleExcelStyles>("Resources\\WriterTemplateExample.xlsx"))
  {
      // Get Excel byte-array
      byte[] excelBytes = writerEngine.Create(new IExcelSheetBuilder<WriterDataContext>[]
      {
          new ExcelSheetBuilderExample()
      },
      CreateDataContext());
  
      // Write to output
      File.WriteAllBytes("WriterTemplateOutput.xlsx", excelBytes);
  }
```

Partial content of ExcelSheetBuilderExample
```C#

public void Build(IExcelSheetWriter sheet, WriterDataContext dataContext)
{
    InsertCustomFormats(sheet, 2);
    InsertDataContextValues(sheet, dataContext, 6);
    InsertFormulas(sheet, dataContext, 9, 8); // Available in source code
    InsertImages(sheet, 18); // Available in source code
}

private void InsertCustomFormats(IExcelSheetWriter sheet, int initialRow)
{
    IRowCursor cursor = sheet.InitializeCursor(new CellReference(initialRow, "A"));
    sheet.InsertValue("This cell is red", ExampleExcelStyles.RedCell);
    cursor.NextColumn();
    sheet.InsertValue("This is green and use Arial font size 13", ExampleExcelStyles.GreenArialCell);
    cursor.NextRowFromOrigin();
    cursor.NextRowFromOrigin();
    sheet.InsertValue("This row has custom height: 40");
    sheet.SetRowHeight(40);
}

private void InsertDataContextValues(IExcelSheetWriter sheet, WriterDataContext dataContext, int initialRow)
{
    // Write DataContext values
    IRowCursor cursor = sheet.InitializeCursor(new CellReference(initialRow, "A"));
    sheet.InsertValue("Now DataContext values: (merge 4 columns and 2 rows)");
    sheet.Merge(4, 2, ExampleExcelStyles.BlueCell);
    cursor.NextRowFromOrigin();
    cursor.NextRowFromOrigin();

    sheet.InsertValue($"{dataContext.Title}: (merge 2 columns)");
    sheet.MergeColumns(2, ExampleExcelStyles.LightBlueCell);

    cursor.NextColumn();
    cursor.NextColumn();
    sheet.InsertValue($"{dataContext.SubTitle}: (merge 2 columns)");
    sheet.MergeColumns(2, ExampleExcelStyles.LightBlueCell);
    cursor.NextRowFromOrigin();

    // User info
    sheet.InsertValue("Name", ExampleExcelStyles.HeaderTableCell);
    cursor.NextColumn();
    sheet.InsertValue("LastName", ExampleExcelStyles.HeaderTableCell);
    cursor.NextColumn();
    sheet.InsertValue("Age", ExampleExcelStyles.HeaderTableCell);
    cursor.NextColumn();
    sheet.InsertValue("Child count", ExampleExcelStyles.HeaderTableCell);
    cursor.NextRowFromOrigin();
    foreach (var user in dataContext.Users)
    {
        sheet.InsertValue(user.Name, ExampleExcelStyles.TableCell);
        cursor.NextColumn();
        sheet.InsertValue(user.LastName, ExampleExcelStyles.TableCell);
        cursor.NextColumn();
        sheet.InsertValue(user.Age, ExampleExcelStyles.TableCell);
        cursor.NextColumn();
        sheet.InsertValue(user.ChildCount, ExampleExcelStyles.TableCell);
        cursor.NextRowFromOrigin();
    }
}
```
   
   
   
   

