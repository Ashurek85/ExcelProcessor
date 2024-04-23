namespace ExcelProcessor.Abstractions.Generator.ReaderResults
{
    public interface IExcelReaderError
    {
        public bool IsGlobalError { get; set; }
        public int? LineNumError { get; set; }
        public string ErrorDescription { get; set; }
        public string Cell { get; set; }
    }
}
