using System;
using System.Collections.Generic;
using Word = Microsoft.Office.Interop.Word;

namespace CoverageKiller2.DOM
{
    public class CKCellRef : ICellRef<CKCell>
    {
        public CKTable Table { get; }
        public IEnumerable<int> CellIndexes { get; }
        public IDOMObject Parent { get; }

        public int WordRow { get; private set; }
        public int WordCol { get; private set; }

        public CKCellRef(Word.Cell wordCell, IDOMObject parent = null)
        {
            Table = CKTable.FromRange(wordCell.Range);
            CellIndexes = new List<int>() { Table.IndexOf(wordCell) };
            WordRow = wordCell.Row.Index;
            WordCol = wordCell.Column.Index;
            Parent = parent ?? Table;
        }
    }

    public class CKCell : CKRange
    {
        public Word.Cell COMCell { get; }
        public CKTable Table { get; }
        public int WordRow { get; }
        public int WordColumn { get; }
        public CKCellRef CellRef { get; }
        public CKCell(CKTable table, IDOMObject parent, Word.Cell wdCell, int wordRow, int wordColumn)
            : base(wdCell.Range, parent)
        {
            Table = table ?? throw new ArgumentNullException(nameof(table));
            COMCell = wdCell ?? throw new ArgumentNullException(nameof(wdCell));
            WordRow = wordRow;
            WordColumn = wordColumn;
            CellRef = new CKCellRef(COMCell, parent);
        }

        public Word.WdColor BackgroundColor
        {
            get => COMCell.Shading.BackgroundPatternColor;
            set => COMCell.Shading.BackgroundPatternColor = value;
        }

        public Word.WdColor ForegroundColor
        {
            get => COMCell.Shading.ForegroundPatternColor;
            set => COMCell.Shading.ForegroundPatternColor = value;
        }
    }
}