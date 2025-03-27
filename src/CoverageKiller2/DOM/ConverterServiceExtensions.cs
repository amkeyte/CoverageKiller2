// CKTable.Converters.cs
using System;
using System.Collections.Generic;
using System.Linq;

namespace CoverageKiller2.DOM
{

    public static class ConverterServiceExtensions
    {
        public static CKCell GetCell(this CKTable.CKCellRefConverterService service, IDOMObject parent, CellRefCoord cellRef)
        {
            if (service == null) throw new ArgumentNullException(nameof(service));
            if (cellRef == null) throw new ArgumentNullException(nameof(cellRef));

            var gridRef = new CKGridCellRef(cellRef.GridX1, cellRef.GridY1, cellRef.GridX2, cellRef.GridY2);
            var grid = service.Grid;

            var masterCell = grid.GetMasterCells(gridRef).FirstOrDefault();
            if (masterCell == null)
                throw new InvalidOperationException("No matching GridCell found.");

            return new CKCell(service.Table, parent, masterCell.COMCell, masterCell.GridRow + 1, masterCell.GridCol + 1);
        }
        public static CKCellsRect GetCells(this CKTable.CKCellRefConverterService converter, CKTable table, IDOMObject parent, ICellRef<CKCellsRect> cellRef)
        {
            if (converter == null) throw new ArgumentNullException(nameof(converter));
            if (table == null) throw new ArgumentNullException(nameof(table));
            if (parent == null) throw new ArgumentNullException(nameof(parent));
            if (cellRef == null) throw new ArgumentNullException(nameof(cellRef));

            var gridRef = new CKGridCellRef(cellRef.GridX1, cellRef.GridY1, cellRef.GridX2, cellRef.GridY2);
            var gridCells = converter.Grid.GetMasterCells(gridRef);

            var cells = new List<CKCell>();
            foreach (var gridCell in gridCells)
            {
                int wordRow = gridCell.GridRow + 1;
                int wordCol = gridCell.GridCol + 1;

                cells.Add(new CKCell(table, parent, gridCell.COMCell, wordRow, wordCol));
            }

            return new CKCellsRect(table, cellRef);
        }

    }

}

