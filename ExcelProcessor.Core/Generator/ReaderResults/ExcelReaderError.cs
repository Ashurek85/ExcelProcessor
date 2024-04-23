using ExcelProcessor.Abstractions.Generator.ReaderResults;

namespace ExcelProcessor.Core.Generator.ReaderResults
{
    public class ExcelReaderError : IExcelReaderError
    {
        public bool IsGlobalError { get; set; }
        public int? LineNumError { get; set; }
        public string ErrorDescription { get; set; }
        public string Cell { get; set; }
    }
}
