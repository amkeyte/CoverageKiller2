using Serilog;
using System;
using System.Linq;
using Word = Microsoft.Office.Interop.Word;

namespace CoverageKiller2
{
    public class CKColumn
    {
        private Word.Column _column;

        // Constructor to initialize cells from a table range
        //that represents a column, ignoring merged cells
        public CKColumn(Word.Column column)
        {
            try
            {
                if (column is null || new CKCells(column.Cells).ContainsMerged)
                    throw new Exception("Shit failed.");

                _column = column;

            }
            catch (Exception ex)
            {
                Log.Error(ex, "Crashing");
                throw ex;
            }
        }
        public CKCells Cells => new CKCells(_column.Cells);
        public bool ContainsMerged => Cells.ContainsMerged;
        // Property to get the index of the column in the table
        public int Index => _column.Cells[1].ColumnIndex;

        // Example: Property to get the width of the column
        public float Width
        {
            get => _column.Width;
            set => _column.Width = value;
        }

        // Example: Property to get the number of cells in the column
        public int CellCount => _column.Cells.Count;

        // Example: Method to select the entire column


        public void Delete()
        {
            Log.Debug("TRACE => {class}.{func}() = {pVal1}",
                nameof(CKColumn),
                nameof(Delete),
                $"{nameof(CKColumn)}[{nameof(Index)} = {Index}]" +
                $"[{nameof(ContainsMerged)} = {ContainsMerged}]" +
                $"[Cell(1) Text = {Cells[1].Text}]");

            if (ContainsMerged)
            {
                DeleteLeavingMerged();
                return;
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
