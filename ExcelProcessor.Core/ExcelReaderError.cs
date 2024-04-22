namespace ExcelProcessor.Core
{
    public class ExcelReaderError
    {
        public bool IsGlobalError { get; set; }
        public int? LineNumError { get; set; }
        public string ErrorDescription { get; set; }
        public string Cell { get; set; }
    }
}
