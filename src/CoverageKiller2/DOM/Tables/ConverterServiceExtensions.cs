// CKTable.Converters.cs
using System;
using System.Linq;
using Word = Microsoft.Office.Interop.Word;

namespace CoverageKiller2.DOM.Tables
{

    //****** NEVER call one conversion methods from another. ******

    public static class ConverterServiceExtensions
    {
        public static CKCellRef GetCellRef(
            this CKTable.CKCellRefConverterService service,
            CKGridCellRef gridRef,
            IDOMObject parent = null)
        {
            var master = service.Grid.GetMasterCells(gridRef).First();
            //return new CKCellRef(master.COMCell, parent ?? service.Table);
            return new CKCellRef(master.COMCell.RowIndex, master.COMCell.ColumnIndex, master.Snapshot, parent);
        }

        public static CKGridCellRef GetGridCellRef(
            this CKTable.CKCellRefConverterService service,
            int index)
        {
            if (service == null) throw new ArgumentNullException(nameof(service));
            if (index < 1) throw new ArgumentOutOfRangeException(nameof(index));

            var wordCell = service.Table.COMTable.Range.Cells[index];
            return new CKGridCellRef(
                wordCell.RowIndex - 1,
                wordCell.ColumnIndex - 1,
                wordCell.RowIndex - 1,
                wordCell.ColumnIndex - 1);
        }


        public static CKGridCellRef GetGridCellRef(
            this CKTable.CKCellRefConverterService service,
            CKCellRef cellRef)
        {
            if (cellRef == null) throw new ArgumentNullException(nameof(cellRef));
            return new CKGridCellRef(
                cellRef.RowIndex - 1,
                cellRef.ColumnIndex - 1,
                cellRef.RowIndex - 1,
                cellRef.ColumnIndex - 1);
        }

        public static CKGridCellRef GetGridCellRef(
            this CKTable.CKCellRefConverterService service,
            Word.Cell wordCell)
        {
            if (wordCell == null) throw new ArgumentNullException(nameof(wordCell));

            return new CKGridCellRef(
                wordCell.RowIndex - 1,
                wordCell.ColumnIndex - 1,
                wordCell.RowIndex - 1,
                wordCell.ColumnIndex - 1);
        }
    }

}

