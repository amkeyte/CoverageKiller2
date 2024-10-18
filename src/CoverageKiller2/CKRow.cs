using System;
using Word = Microsoft.Office.Interop.Word;
namespace CoverageKiller2
{
    public class CKRow
    {
        private Word.Row _row;

        // Constructor to initialize CKRow with a Word.Row
        public CKRow(Word.Row row)
        {
            _row = row ?? throw new ArgumentNullException(nameof(row));
        }
        public CKCells Cells => new CKCells(_row.Cells);
        public bool ContainsMerged => Cells.ContainsMerged;
        // Property to access row's index
        public int Index => _row.Index;

        // Property to get the number of cells in the row
        public int CellCount => _row.Cells.Count;

        // Access a cell by its index (1-based index in Word)
        public CKCell GetCell(int index)
        {
            if (index < 1 || index > _row.Cells.Count)
                throw new ArgumentOutOfRangeException(nameof(index), "Index must be between 1 and CellCount.");

            return new CKCell(_row.Cells[index]);
        }

        // Set or get the row's height
        public float Height
        {
            get => _row.Height;
            set => _row.Height = value;
        }

        // Set or get the row's height rule (auto, at least, or exact)
        public Word.WdRowHeightRule HeightRule
        {
            get => _row.HeightRule;
            set => _row.HeightRule = value;
        }

        // Property to get or set row shading
        public Word.WdColor ShadingBackgroundColor
        {
            get => (Word.WdColor)_row.Shading.BackgroundPatternColor;
            set => _row.Shading.BackgroundPatternColor = (Word.WdColor)value;
        }

        // Merges all cells in the row
        public void Merge()
        {
            _row.Cells.Merge();
        }

        // Deletes the row
        public void Delete()
        {
            _row.Delete();
        }

        // Selects the row
        public void Select()
        {
            _row.Select();
        }
    }
}


