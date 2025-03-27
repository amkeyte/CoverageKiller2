using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using Word = Microsoft.Office.Interop.Word;

namespace CoverageKiller2.DOM
{
    /// <summary>
    /// Represents an arbitrary collection of CKCell objects.
    /// </summary>
    public abstract class CKCells : IEnumerable<CKCell>, IDOMObject
    {
        protected List<CKCell> _cells;
        public CKTable Table { get; protected set; }
        public ICellRef<IDOMObject> CellRef { get; protected set; }

        protected CKCells(CKTable table, ICellRef<CKCells> cellReference)
        {
            Table = table ?? throw new ArgumentNullException(nameof(table));
            CellRef = cellReference ?? throw new ArgumentNullException(nameof(cellReference));
            _cells = BuildCells().ToList();
        }

        /// <summary>
        /// Derived classes implement this method to build the cell collection.
        /// </summary>
        /// <returns>An enumerable of CKCell objects.</returns>
        protected abstract IEnumerable<CKCell> BuildCells();

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


    /// <summary>

    ///Broken do not use    
    /// </summary>
    //public class CKCellsLinear : CKCells
    //{
    //    public CKCellsLinear(CKTable table, ICellRef<CKCellsLinear> cellReference)
    //        : base(table, cellReference)
    //    {
    //    }

    //    protected override IEnumerable<CKCell> BuildCells()
    //    {
    //        var cellsRect = Table.Converters.GetCells(Table, this, (ICellRef<CKCellsLinear>)CellRef);
    //        return cellsRect;
    //    }

    //}

    /// <summary>
    /// Represents a collection of CKCell objects that form a contiguous rectangular grid.
    /// </summary>
    public class CKCellsRect : CKCells
    {
        public CKCellsRect(CKTable table, ICellRef<CKCellsRect> cellReference)
            : base(table, cellReference)
        {
        }

        protected override IEnumerable<CKCell> BuildCells()
        {
            var cellsRect = Table.Converters.GetCells(Table, this, (ICellRef<CKCellsRect>)CellRef);
            return cellsRect;
        }
    }

}

