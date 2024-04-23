﻿using ExcelProcessor.Abstractions.Pointers;

namespace ExcelProcessor.Abstractions.Generator.ReaderResults
{
    public interface IExcelReaderResult<TEntityReaded>
        where TEntityReaded : class
    {
        IEnumerable<IExcelReaderError> Errors { get; }

        public TEntityReaded EntityReaded { get; }

        void AddGlobalError(string error);

        void AddCellError(string error, ICellReference cellRef);

        void AddLineError(string error, int numLine);

        void Accumulate(IExcelReaderResult<TEntityReaded> otherResults);
    }
}