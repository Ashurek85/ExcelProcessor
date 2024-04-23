# ExcelProcessor

Wrapper over OpenXML for ease of use. Offers the functionality to read and generate Excel files

## Positioning abstractions
Position and movement through Excel is done through two abstractions:
- ICellReference. Defines a position: row and column
- IRowCursor. Cursor with the ability to change cell reference position. The first position is marked as origin
  - NextColumn(): Move to the next column
  - NextRowFromOrigin(): Allows positioning in the next row and column of origin

## Write operations
The entry point for writing in Excel is creating an instance of ExcelGenerator. It has methods to write in excel from:
- Excel Template file
- Byte array that represents a Excel file
- Empty file
