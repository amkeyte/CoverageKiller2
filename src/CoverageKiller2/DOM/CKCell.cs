using Word = Microsoft.Office.Interop.Word;

namespace CoverageKiller2.DOM
{
    public class CKCell : CKRange
    {
        /// <summary>
        /// Avoid use if possible. Probably be hidden.
        /// </summary>
        public Word.Cell COMCell { get; private set; }


        //if cellref is more than one cell, the top left is given. (Consistent with Range.Tables)
        public CKCell(CKCells cellsCollection, CKCellRefLinear cellRef, int Index) :
            this(cellsCollection.Table,
                cellsCollection,
                cellsCollection.Table.WordCell(xxx),
                cellRef.X1,
                cellRef.Y1)
        {

        }

        public CKCell(CKTable table, CKGridCellRefRect cellRef) :
            this(table, table, table.WordCell(cellRef), cellRef.X1, cellRef.Y1)
        {

        }
        // full constructor that requires all parameters.
        private CKCell(CKTable table, IDOMObject parent, Word.Cell wdCell, int row, int col) :
            base(wdCell.Range, parent)
        {
            Table = table;
            CellRef = CKGridCellRefRect.ForCell(row, col);
        }

        // Property to get or set the background color for the cell
        public Word.WdColor BackgroundColor
        {
            get => COMCell.Shading.BackgroundPatternColor;
            set => COMCell.Shading.BackgroundPatternColor = value;
        }

        // Property to get or set the foreground (pattern) color for the cell
        public Word.WdColor ForegroundColor
        {
            get => COMCell.Shading.ForegroundPatternColor;
            set => COMCell.Shading.ForegroundPatternColor = value;
        }

        //Table is not nessisarily the parent!!
        public CKTable Table { get; private set; }
        public CKGridCellRefRect CellRef { get; }
    }
}

