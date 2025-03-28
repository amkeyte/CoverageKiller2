using CoverageKiller2.DOM;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using Word = Microsoft.Office.Interop.Word;

namespace CoverageKiller2.DOM
{
    public class CKCells : IEnumerable<CKCell>, IDOMObject
    {
        protected List<CKCell> _cells = new List<CKCell>();
        public CKTable Table { get; protected set; }
        public ICellRef<CKCells> CellRef { get; protected set; }

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


        public static CKCells FromRef(CKTable table, ICellRef<CKCells> cellRef)
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

