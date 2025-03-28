// CKTable.Converters.cs
using System;
using System.Linq;
using Word = Microsoft.Office.Interop.Word;

namespace CoverageKiller2.DOM
{

    public static class ConverterServiceExtensions
    {
        public static CKCellRef GetCellRef(
            this CKTable.CKCellRefConverterService service,
            CKGridCellRef gridRef,
            IDOMObject parent = null)
        {
            var grid = service.Grid;
            var master = grid.GetMasterCells(gridRef).FirstOrDefault();
            if (master == null)
                throw new ArgumentException("No master cell found in specified grid region.", nameof(gridRef));

            return new CKCellRef(master.COMCell, parent ?? service.Table);
        }

        public static CKGridCellRef GetGridCellRef(
            this CKTable.CKCellRefConverterService service,
            int index)
        {
            if (service == null) throw new ArgumentNullException(nameof(service));
            if (index < 1) throw new ArgumentOutOfRangeException(nameof(index));

            var wordCell = service.Table.COMTable.Range.Cells[index];
            return service.GetGridCellRef(wordCell);
        }


        public static CKGridCellRef GetGridCellRef(
            this CKTable.CKCellRefConverterService service,
            ICellRef<CKCell> cellRef)
        {
            if (cellRef == null) throw new ArgumentNullException(nameof(cellRef));
            return service.GetGridCellRef(
                service.Table.COMTable.Cell(
                    (cellRef as CKCellRef)?.WordRow ?? 1,
                    (cellRef as CKCellRef)?.WordCol ?? 1
                )
            );
        }

        public static CKGridCellRef GetGridCellRef(
            this CKTable.CKCellRefConverterService service,
            Word.Cell cellRef)
        {
            if (cellRef == null) throw new ArgumentNullException(nameof(cellRef));

            var master = service.Grid.GetMasterCells()
                .FirstOrDefault(g => g.COMCell == cellRef);

            if (master == null)
                throw new ArgumentException("Cell is not a master cell in the grid.", nameof(cellRef));

            return new CKGridCellRef(master.GridCol, master.GridRow, master.GridCol, master.GridRow);
        }
    }

}

