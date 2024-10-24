using System;
using Word = Microsoft.Office.Interop.Word;

namespace CoverageKiller2
{
    public class CKCell
    {
        internal Word.Cell COMObject { get; private set; }

        public CKCells Parent { get; private set; }
        internal static CKCell Create(CKCells parent, int index)
        {
            return new CKCell(parent, index);
        }
        private CKCell(CKCells parent, int index)
        {
            Parent = parent;
            COMObject = Parent.COMObject[index];
        }




        // Property to get or set the text in a cell
        public string Text
        {
            //really trimend? maybe let the consumer handle that problem.
            get => COMObject.Range.Text.TrimEnd('\r');
            set => COMObject.Range.Text = value;
        }

        // Property to get or set the background color for the cell
        public Word.WdColor BackgroundColor
        {
            get => (Word.WdColor)COMObject.Shading.BackgroundPatternColor;
            set => COMObject.Shading.BackgroundPatternColor = value;
        }

        // Property to get or set the foreground (pattern) color for the cell
        public Word.WdColor ForegroundColor
        {
            get => (Word.WdColor)COMObject.Shading.ForegroundPatternColor;
            set => COMObject.Shading.ForegroundPatternColor = value;
        }

        // Merges this cell with others (for example, in a row or column)
        public void Merge(CKCell otherCell)
        {
            if (otherCell == null)
                throw new ArgumentNullException(nameof(otherCell));

            COMObject.Merge(otherCell.COMObject);
        }

        // Deletes the cell
        public void Delete()
        {
            COMObject.Delete();
        }

        // Selects the cell in the document
        public void Select()
        {
            COMObject.Select();
        }

        // Property to get the index of the cell in the row
        public int ColumnIndex => COMObject.ColumnIndex;
        // Override Equals to compare by DOM object (Range)
        public override bool Equals(object obj)
        {
            if (obj is CKCell other)
            {
                return this.COMObject.Range.Equals(other.COMObject.Range);
            }
            return false;
        }

        // Override GetHashCode to provide a hash code consistent with Equals
        public override int GetHashCode()
        {
            return COMObject.Range.GetHashCode();
        }



        internal static CKCell Create(CKTable parent, int row, int column)
        {
            var cellRow = CKCells.Create(parent.Rows[row]);
            return CKCell.Create(cellRow, column);
        }

        public int RowIndex => COMObject.RowIndex;
    }

    internal static partial class CKCellExtensions
    {
        //// Extension method to get the cell above
        //public static Word.Cell Up(this Word.Cell cell)
        //{
        //    if (cell == null)
        //        throw new ArgumentNullException(nameof(cell));

        //    try
        //    {
        //        return cell.Range.Tables[1].Cell(cell.Row.Index - 1, cell.Column.Index);
        //    }
        //    catch (ArgumentException)
        //    {
        //        return null; // Return null if the cell above doesn't exist
        //    }
        //}



        //// Extension method to get the cell below
        //public static Word.Cell Down(this Word.Cell cell)
        //{
        //    if (cell == null)
        //        throw new ArgumentNullException(nameof(cell));

        //    try
        //    {
        //        return cell.Range.Tables[1].Cell(cell.Row.Index + 1, cell.Column.Index);
        //    }
        //    catch (ArgumentException)
        //    {
        //        return null; // Return null if the cell below doesn't exist
        //    }
        //}

        //// Extension method to get the cell to the left
        //public static Word.Cell Left(this Word.Cell cell)
        //{
        //    if (cell == null)
        //        throw new ArgumentNullException(nameof(cell));

        //    try
        //    {
        //        return cell.Range.Tables[1].Cell(cell.Row.Index, cell.Column.Index - 1);
        //    }
        //    catch (ArgumentException)
        //    {
        //        return null; // Return null if the cell to the left doesn't exist
        //    }
        //}

        //// Extension method to get the cell to the right
        //public static Word.Cell Right(this Word.Cell cell)
        //{
        //    if (cell == null)
        //        throw new ArgumentNullException(nameof(cell));

        //    try
        //    {
        //        return cell.Range.Tables[1].Cell(cell.Row.Index, cell.Column.Index + 1);
        //    }
        //    catch (ArgumentException)
        //    {
        //        return null; // Return null if the cell to the right doesn't exist
        //    }
        //}

        //public static Word.Table Table(this Word.Cell cell)
        //{
        //    // Check if the cell is null
        //    if (cell == null)
        //    {
        //        throw new ArgumentNullException(nameof(cell));
        //    }

        //    // Access the parent Range and get the Table from that Range
        //    Word.Range cellRange = cell.Range;
        //    return cellRange.Tables.Count > 0 ? cellRange.Tables[1] : null; // Return the first table if it exists
        //}

        //public static bool IsMerged(this Word.Cell cell)
        //{
        //    Log.Debug(@"TRACE => EXTENSION:{class}.{func}({param1})",
        //        nameof(Word.Cell),
        //        nameof(IsMerged),
        //        $"{nameof(cell)} = Cell[{cell.RowIndex},{cell.ColumnIndex}]");
        //    try
        //    {
        //        if (cell != null)
        //        {
        //            // Get the table that contains the cell
        //            Word.Table table = cell.Table();

        //            // Check if the table is not null
        //            if (table != null)
        //            {
        //                // attempt to use rows or columns. It will fail if there is a merge.
        //                _ = table.Rows[1];
        //                _ = table.Columns[1];


        //            }
        //        }
        //    }
        //    catch (System.Runtime.InteropServices.COMException ex)
        //    {
        //        // Check for specific error codes for merged cells
        //        if (ex.Message.Contains("column") == -2146827284) // Error code for vertically merged cells (5991)
        //        {
        //            Log.Debug(@"TRACE => EXTENSION:{class}.{func}({param1}) {message}",
        //                nameof(Word.Cell),
        //                nameof(IsMerged),
        //                string.Empty,
        //                "There was a vertiacally merged cell found in the table");

        //            return true;
        //        }
        //        else if (ex.ErrorCode == -2146827283) // Error code for horizontally merged cells (5992)
        //        {
        //            Log.Debug(@"TRACE => EXTENSION:{class}.{func}({param1}) {message}",
        //                nameof(Word.Cell),
        //                nameof(IsMerged),
        //                string.Empty,
        //                "There was a horizontally merged cell found in the table");
        //            // Optionally log the event
        //            // Log.Information("Cell [{0},{1}] has horizontally merged cells.", cell.RowIndex, cell.ColumnIndex);
        //            return true;
        //        }
        //        else
        //        {
        //            // Optionally log the generic error
        //            // Log.Error("Error {0}: {1} in cell [{2},{3}]", ex.ErrorCode, ex.Message, cell.RowIndex, cell.ColumnIndex);
        //            throw; // Rethrow unexpected exceptions
        //        }
        //    }

        //    return false; // Not merged
        //}


        //Log.Debug(@"TRACE => EXTENSION:{class}.{func}({param1}) ::{internals}:: = {return}",
        //    nameof(Cell),
        //    nameof(IsMerged),
        //    $"{nameof(_cell)} = Cell[{_cell.RowIndex},{_cell.ColumnIndex}]",
        //    $"{nameof(isHorizontallyMerged)} = {isHorizontallyMerged} || " +
        //        $"{nameof(isVerticallyMerged)} = {isVerticallyMerged}",
        //    isHorizontallyMerged || isVerticallyMerged);

        // Helper method to get the cell directly below (Word.Cell does not have a Down property)
        //private static Word.Cell GetCellBelow(Word.Cell cell)
        //{
        //    try
        //    {
        //        Word.Row nextRow = cell.Row.Next;
        //        if (nextRow != null)
        //        {
        //            return nextRow.Cells[cell.ColumnIndex];
        //        }
        //    }
        //    catch
        //    {
        //        // Handle any exceptions if accessing the next row or cell fails
        //    }
        //    return null;
        //}




        //public static Word.Range MergeArea(this Word.Cell cell)
        //{
        //    // If the cell is not merged, return its own range
        //    if (!cell.IsMerged())
        //    {
        //        return cell.Range;
        //    }

        //    // Create a Range that starts as the current cell's range
        //    Word.Range mergedRange = cell.Range.Duplicate;

        //    // Check for merged cells using Up, Down, Left, and Right properties
        //    // Check upwards
        //    MergeCellsInDirection(cell, ref mergedRange, cell.Up, true);

        //    // Check downwards
        //    MergeCellsInDirection(cell, ref mergedRange, cell.Down, false);

        //    // Check left
        //    MergeCellsInDirection(cell, ref mergedRange, cell.Left, true, false);

        //    // Check right
        //    MergeCellsInDirection(cell, ref mergedRange, cell.Right, true, true);

        //    return mergedRange;
        //}

        //private static void MergeCellsInDirection(Word.Cell startingCell, ref Word.Range mergedRange,
        //    Func<Word.Cell> getNextCell, bool isVertical, bool moveRight = true)
        //{
        //    Word.Cell nextCell = getNextCell();
        //    while (nextCell != null && nextCell.IsMerged() &&
        //           nextCell.Range.Start == startingCell.Range.Start &&
        //           nextCell.Range.End == startingCell.Range.End)
        //    {
        //        // Extend the range to include the next cell
        //        mergedRange.Start = Math.Min(mergedRange.Start, nextCell.Range.Start);
        //        mergedRange.End = Math.Max(mergedRange.End, nextCell.Range.End);

        //        nextCell = isVertical ? (moveRight ? nextCell.Down() : nextCell.Up()) : (moveRight ? nextCell.Right() : nextCell.Left());
        //    }
        //}
    }
}

