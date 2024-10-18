using System;
using System.Linq;
using Word = Microsoft.Office.Interop.Word;

namespace CoverageKiller2
{
    public class CKColumn
    {
        private Word.Column _column;

        // Constructor to initialize CKColumn with a Word.Column
        public CKColumn(Word.Column column)
        {
            _column = column ?? throw new ArgumentNullException(nameof(column));
        }
        public CKCells Cells => new CKCells(_column.Cells);
        public bool ContainsMerged => Cells.ContainsMerged;
        // Property to get the index of the column in the table
        public int Index => _column.Index;

        // Example: Property to get the width of the column
        public float Width
        {
            get => _column.Width;
            set => _column.Width = value;
        }

        // Example: Property to get the number of cells in the column
        public int CellCount => _column.Cells.Count;

        // Example: Method to select the entire column
        public void Select()
        {
            _column.Select();
        }

        public void Delete()
        {
            if (ContainsMerged)
            {
                DeleteLeavingMerged();
            }

            _column.Delete();
        }

        private void DeleteLeavingMerged()
        {
            //if it doesn't work, examine how the merged cells effect the coordinates.
            //Ideally accessing the col.cells index acts expeded for referencing the cell,
            //but using the merged cell's ColumnIndex is possibly not the same.
            Cells.Where(c => !c.IsMerged).ToList()
                .ForEach(c => c.Delete());
        }
    }
}
