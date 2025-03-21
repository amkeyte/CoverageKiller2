using System;
using Word = Microsoft.Office.Interop.Word;

namespace CoverageKiller2.DOM
{
    public class CKCell : CKRange
    {
        /// <summary>
        /// Avoid use if possible. Probably be hidden.
        /// </summary>
        internal Word.Cell COMCell { get; private set; }


        /// <summary>
        /// CELLS ARE BROKEN DO NOT USE!!!!
        /// </summary>
        /// <param name="parent"></param>
        /// <param name="index"></param>
        /// <exception cref="ArgumentNullException"></exception>
        private CKCell(CKCells parent, int index) : base(parent.COMObject[1].Range)
        {

            //Parent = parent;

            //if (Parent.COMObject is null) throw new ArgumentNullException("parent");

            //COMObject = Parent.COMObject[index];


        }
        public CKCell(Word.Cell cell) : base(cell.Range)
        {
            COMCell = cell;
        }



        // Property to get or set the text in a cell
        public string Text
        {
            //really trimend? maybe let the consumer handle that problem.
            get => COMRange.Text.TrimEnd('\r');
            set => COMRange.Text = value;
        }

        // Property to get or set the background color for the cell
        public Word.WdColor BackgroundColor
        {
            get => (Word.WdColor)COMCell.Shading.BackgroundPatternColor;
            set => COMCell.Shading.BackgroundPatternColor = value;
        }

        // Property to get or set the foreground (pattern) color for the cell
        public Word.WdColor ForegroundColor
        {
            get => (Word.WdColor)COMCell.Shading.ForegroundPatternColor;
            set => COMCell.Shading.ForegroundPatternColor = value;
        }

        // Merges this cell with others (for example, in a row or column)
        public void Merge(CKCell otherCell)
        {
            if (otherCell == null)
                throw new ArgumentNullException(nameof(otherCell));

            COMCell.Merge(otherCell.COMCell);
        }

        // Deletes the cell
        public void Delete()
        {
            COMCell.Delete();
        }

        // Selects the cell in the document
        public void Select()
        {
            COMCell.Select();
        }

        // Property to get the index of the cell in the row
        public int ColumnIndex => COMCell.ColumnIndex;
        // Override Equals to compare by DOM object (Range)
        public override bool Equals(object obj)
        {
            if (obj is CKCell other)
            {
                return this.COMRange.Equals(other.COMCell.Range);
            }
            return false;
        }

        // Override GetHashCode to provide a hash code consistent with Equals
        public override int GetHashCode()
        {
            return COMCell.Range.GetHashCode();
        }



        //internal static CKCell Create(CKTable parent, int row, int column)
        //{
        //    var cellRow = CKCells.Create(parent.Rows[row]);
        //    return CKCell.Create(cellRow, column);
        //}

        public int RowIndex => COMCell.RowIndex;
    }
}

