// CKTable.Converters.cs
using System;
using System.Linq;
using Word = Microsoft.Office.Interop.Word;

namespace CoverageKiller2.DOM
{
    //NEVER cal other conversion methods.
    public static class ConverterServiceExtensions
    {
        public static CKCellRef GetCellRef(
            this CKTable.CKCellRefConverterService service,
            CKGridCellRef gridRef,
            IDOMObject parent = null)
        {
            var master = service.Grid.GetMasterCells(gridRef).First();
            return new CKCellRef(master.COMCell, parent ?? service.Table);
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
                cellRef.WordRow - 1,
                cellRef.WordCol - 1,
                cellRef.WordRow - 1,
                cellRef.WordCol - 1);
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

