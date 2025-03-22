using Word = Microsoft.Office.Interop.Word;

namespace CoverageKiller2.DOM
{
    public class CKCell : CKRange
    {
        /// <summary>
        /// Avoid use if possible. Probably be hidden.
        /// </summary>
        public Word.Cell COMCell { get; private set; }


        public CKCell(Word.Cell cell) : base(cell.Range)
        {
            COMCell = cell;
        }



        // Property to get or set the text in a cell
        public string Text
        {
            get => COMRange.Text;
            set => COMRange.Text = value;
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
    }
}

