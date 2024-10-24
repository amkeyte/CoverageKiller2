using Serilog;
using Word = Microsoft.Office.Interop.Word;

namespace CoverageKiller2
{
    public class CKColumn
    {
        internal Word.Column COMObject { get; private set; }



        public CKColumns Parent { get; private set; }
        // Constructor to initialize cells from a table range
        //that represents a column, ignoring merged cells
        private CKColumn(CKColumns parent, int index)
        {
            Parent = parent;
            COMObject = Parent.COMObject[index];

            //
            //    //if (column is null || new CKCells(column.Cells).ContainsMerged)
            //    //    throw new Exception("Shit failed.");

            //    _column = column;

            //}
            //catch (Exception ex)
            //{
            //    Log.Error(ex, "Crashing");
            //    throw ex;
            //}
        }
        public CKCells Cells => CKCells.Create(this);// new CKCells(COMObject.Cells);
        //public bool ContainsMerged => Cells.ContainsMerged;
        // Property to get the index of the column in the table
        public int Index => COMObject.Index;

        // Example: Property to get the width of the column
        public float Width
        {
            get => COMObject.Width;
            set => COMObject.Width = value;
        }



        public static CKColumn Create(CKColumns parent, int index)
        {
            return new CKColumn(parent, index);
        }

        // Example: Method to select the entire column


        public void Delete()
        {
            Log.Debug("TRACE => {class}.{func}() = {pVal1}",
                nameof(CKColumn),
                nameof(Delete),
                $"{nameof(CKColumn)}[{nameof(Index)} = {Index}]" +
                //$"[{nameof(ContainsMerged)} = {ContainsMerged}]" +
                $"[Cell(1) Text = {Cells[1].Text}]");

            //if (ContainsMerged)
            //{
            //    DeleteLeavingMerged();
            //    return;
            //}

            COMObject.Delete();
        }

        //private void DeleteLeavingMerged()
        //{
        //    //if it doesn't work, examine how the merged cells effect the coordinates.
        //    //Ideally accessing the col.cells index acts expeded for referencing the cell,
        //    //but using the merged cell's ColumnIndex is possibly not the same.
        //    Cells.Where(c => !c.IsMerged).ToList()
        //        .ForEach(c => c.Delete());
        //}
    }
}
