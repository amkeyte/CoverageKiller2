using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using Word = Microsoft.Office.Interop.Word;

namespace CoverageKiller2.DOM.Tables
{
    public interface ICellRef<out T> where T : IDOMObject
    {
        CKTable Table { get; }
        IEnumerable<int> CellIndexes { get; }
        IDOMObject Parent { get; }
    }
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
    public class CKCells : IEnumerable<CKCell>, IDOMObject
    {
        protected List<CKCell> _cells = new List<CKCell>();
        public CKTable Table { get; protected set; }
        public CKCellRef CellRef { get; protected set; }

        protected CKCells() { }

        protected virtual IEnumerable<CKCell> BuildCells()
        {
            if (CellRef == null || CellRef.CellIndexes == null)
                throw new InvalidOperationException("CellRef or CellIndexes is null.");

            foreach (var i in CellRef.CellIndexes)
            {
                var gcr = Table.Converters.GetGridCellRef(i);
                var cellRef = Table.Converters.GetCellRef(gcr);
                _cells.Add(Table.Cell(cellRef));
            }

            return _cells;
        }


        public static CKCells FromRef(CKTable table, CKCellRef cellRef)
        {
            if (table == null || cellRef == null)
                throw new ArgumentNullException();

            var instance = new CKCells
            {
                Table = table,
                CellRef = cellRef
            };

            instance._cells = instance.BuildCells().ToList();
            return instance;
        }

        public int Count => _cells.Count;
        public CKDocument Document => Table.Document;
        public Word.Application Application => Table.Application;
        public IDOMObject Parent => Table;
        public bool IsDirty => Table.IsDirty || _cells.Any(c => c.IsDirty);
        public bool IsOrphan => Document.IsOrphan;

        public CKCell this[int index]
        {
            get
            {
                if (index < 1 || index > Count)
                    throw new ArgumentOutOfRangeException(nameof(index));
                return _cells[index - 1];
            }
        }

        public IEnumerator<CKCell> GetEnumerator() => _cells.GetEnumerator();
        IEnumerator IEnumerable.GetEnumerator() => GetEnumerator();
    }

}