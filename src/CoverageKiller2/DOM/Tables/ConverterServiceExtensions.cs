// CKTable.Converters.cs
using CoverageKiller2.Logging;
using System;
using Word = Microsoft.Office.Interop.Word;

namespace CoverageKiller2.DOM.Tables
{

    //****** NEVER call one conversion methods from another. ******

    public static class ConverterServiceExtensions
    {
        public static CKCellRef GetCellRef(
            this CKTable.CKCellRefConverterService service,
            CKGridCellRef gridRef,
            IDOMObject parent)
        {
            var master = service.Grid.GetMasterCell(gridRef)
                ?? throw new CKDebugException("No master cell returned.");
            //.First();
            return new CKCellRef(master.GridRow, master.GridCol, service.Table, parent);
        }

        public static CKGridCellRef GetGridCellRef(
            this CKTable.CKCellRefConverterService service,
            int index)
        {
            if (service == null) throw new ArgumentNullException(nameof(service));
            if (index < 1) throw new ArgumentOutOfRangeException(nameof(index));

            var wordCell = service.Table.COMTable.Range.Cells[index];
            return new CKGridCellRef(
                wordCell.RowIndex,
                wordCell.ColumnIndex,
                wordCell.RowIndex,
                wordCell.ColumnIndex);
        }


        public static CKGridCellRef GetGridCellRef(
            this CKTable.CKCellRefConverterService service,
            CKCellRef cellRef)
        {
            //LH.Debug("Tracker[!sd]");
            if (cellRef == null) throw new ArgumentNullException(nameof(cellRef));

            return new CKGridCellRef(
                cellRef.RowIndex,
                cellRef.ColumnIndex,
                cellRef.RowIndex,
                cellRef.ColumnIndex);
        }
        public static CKGridCellRef GetGridCellRef(
            this CKTable.CKCellRefConverterService service,
            CKRowCellRef rowRef)
        {
            if (rowRef == null) throw new ArgumentNullException(nameof(rowRef));
            return new CKGridCellRef(
                rowRef.Index, 1, rowRef.Index, service.Grid.ColCount);

        }


        public static CKGridCellRef GetGridCellRef(
            this CKTable.CKCellRefConverterService service,
            Word.Cell wordCell)
        {
            if (wordCell == null) throw new ArgumentNullException(nameof(wordCell));

            return new CKGridCellRef(
                wordCell.RowIndex,
                wordCell.ColumnIndex,
                wordCell.RowIndex,
                wordCell.ColumnIndex);
        }
    }

}

