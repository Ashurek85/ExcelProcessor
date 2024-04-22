using ExcelProcessor.Core.Pointers;

namespace ExcelProcessor.Core
{
    public class ExcelReadResult<TEntity>
        where TEntity : class, new()
    {
        private readonly List<ExcelReaderError> errors = new List<ExcelReaderError>();

        public TEntity Entity { get; set; } = new TEntity();
        public IEnumerable<ExcelReaderError> Errors
        {
            get => errors;
        }

        public bool HasErrors
        {
            get => errors != null && errors.Any();
        }

        public void AddGlobalError(string error)
        {
            errors.Add(new ExcelReaderError()
            {
                IsGlobalError = true,
                ErrorDescription = error,
            });
        }

        public void AddCellError(string error, CellReference cellRef)
        {
            errors.Add(new ExcelReaderError()
            {
                LineNumError = cellRef.Row,
                ErrorDescription = error,
                Cell = cellRef?.ToExcelString()
            });
        }

        public void AddLineError(string error, int numLine)
        {
            errors.Add(new ExcelReaderError()
            {
                LineNumError = numLine,
                ErrorDescription = error,
            });
        }

        public void Accumulate(ExcelReadResult<TEntity> otherResults)
        {
            if (otherResults != null)
            {
                Entity = otherResults.Entity;
                errors.AddRange(otherResults.Errors);
            }
        }
    }
}
