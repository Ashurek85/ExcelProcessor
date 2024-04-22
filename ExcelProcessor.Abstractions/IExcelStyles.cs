namespace ExcelProcessor.Abstractions
{
    using DocumentFormat.OpenXml.Spreadsheet;
    public interface IExcelStyles
    {
        void Inyect(Stylesheet stylesSheet);

        uint GetStyleId(string styleName);
    }
}
