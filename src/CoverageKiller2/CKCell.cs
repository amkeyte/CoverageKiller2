using System;
using System.Linq;
using Word = Microsoft.Office.Interop.Word;

namespace CoverageKiller2
{
    public class CKCell
    {
        private Word.Cell _cell;

        public CKCell(Word.Cell cell)
        {
            _cell = cell ?? throw new ArgumentNullException(nameof(cell));
        }

        public bool IsMerged => _cell.IsMerged();




        // Property to get or set the text in a cell
        public string Text
        {
            get => _cell.Range.Text.TrimEnd('\r');
            set => _cell.Range.Text = value;
        }

        // Property to get or set the background color for the cell
        public Word.WdColor BackgroundColor
        {
            get => (Word.WdColor)_cell.Shading.BackgroundPatternColor;
            set => _cell.Shading.BackgroundPatternColor = value;
        }

        // Property to get or set the foreground (pattern) color for the cell
        public Word.WdColor ForegroundColor
        {
            get => (Word.WdColor)_cell.Shading.ForegroundPatternColor;
            set => _cell.Shading.ForegroundPatternColor = value;
        }

        // Merges this cell with others (for example, in a row or column)
        public void Merge(CKCell otherCell)
        {
            if (otherCell == null)
                throw new ArgumentNullException(nameof(otherCell));

            _cell.Merge(otherCell._cell);
        }

        // Deletes the cell
        public void Delete()
        {
            _cell.Delete();
        }

        // Selects the cell in the document
        public void Select()
        {
            _cell.Select();
        }

        // Property to get the index of the cell in the row
        public int Index => _cell.ColumnIndex;
        // Override Equals to compare by DOM object (Range)
        public override bool Equals(object obj)
        {
            if (obj is CKCell other)
            {
                return this._cell.Range.Equals(other._cell.Range);
            }
            return false;
        }

        // Override GetHashCode to provide a hash code consistent with Equals
        public override int GetHashCode()
        {
            return _cell.Range.GetHashCode();
        }
    }
    internal static partial class CKCellExtensions
    {
        // Extension method to get the cell above
        public static Word.Cell Up(this Word.Cell cell)
        {
            if (cell == null)
                throw new ArgumentNullException(nameof(cell));

            try
            {
                return cell.Range.Tables[1].Cell(cell.Row.Index - 1, cell.Column.Index);
            }
            catch (ArgumentException)
            {
                return null; // Return null if the cell above doesn't exist
            }
        }

        public static bool IsMerged(this Word.Cell _cell)
        {
            return new[] { _cell.Next, _cell.Down() }
                .Any(neighbor => neighbor != null &&
                    _cell.Range.Start == neighbor.Range.Start &&
                    _cell.Range.End == neighbor.Range.End);
        }

        // Extension method to get the cell below
        public static Word.Cell Down(this Word.Cell cell)
        {
            if (cell == null)
                throw new ArgumentNullException(nameof(cell));

            try
            {
                return cell.Range.Tables[1].Cell(cell.Row.Index + 1, cell.Column.Index);
            }
            catch (ArgumentException)
            {
                return null; // Return null if the cell below doesn't exist
            }
        }

        // Extension method to get the cell to the left
        public static Word.Cell Left(this Word.Cell cell)
        {
            if (cell == null)
                throw new ArgumentNullException(nameof(cell));

            try
            {
                return cell.Range.Tables[1].Cell(cell.Row.Index, cell.Column.Index - 1);
            }
            catch (ArgumentException)
            {
                return null; // Return null if the cell to the left doesn't exist
            }
        }

        // Extension method to get the cell to the right
        public static Word.Cell Right(this Word.Cell cell)
        {
            if (cell == null)
                throw new ArgumentNullException(nameof(cell));

            try
            {
                return cell.Range.Tables[1].Cell(cell.Row.Index, cell.Column.Index + 1);
            }
            catch (ArgumentException)
            {
                return null; // Return null if the cell to the right doesn't exist
            }
        }
    }
}
