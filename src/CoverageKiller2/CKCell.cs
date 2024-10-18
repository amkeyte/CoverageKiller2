using System;
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
}
