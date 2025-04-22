using CoverageKiller2.Logging;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;

namespace CoverageKiller2.DOM.Tables
{
    public class CKRowCellRef : CKCellRef, ICellRef<CKRow>
    {
        public CKRowCellRef(int rowIndex, CKTable table, IDOMObject parent) :
            base(rowIndex, table.Columns.Count(), table, parent)
        {
            this.Ping();
            if (parent == null) throw new ArgumentNullException(nameof(parent));
            if (table == null) throw new ArgumentNullException(nameof(table));
            Index = rowIndex;
            Table = table;
            Parent = parent;
            this.Pong();
        }
        /// <inheritdoc/>
        public IDOMObject Parent { get; }
        ///<summary>
        /// Gets the one-based Word row index of the referenced cell.
        /// </summary>
        public int Index { get; }
        public CKTable Table { get; }
    }

    public class CKRow : CKCells
    {
        public override CKCell this[int index]
        {
            get
            {
                if (index < 1 || index > Count)
                    throw new ArgumentOutOfRangeException(nameof(index));

                return CellsList_1[index - 1];
            }
        }
        public CKRow(Base1List<CKCell> cells_1, CKRowCellRef cellRef, IDOMObject parent) : base(cells_1, cellRef.Parent)
        {
            Parent = parent;
            CellRef = cellRef;
        }
        public CKRow(CKRowCellRef rowRef, IDOMObject parent) :
            base(parent)
        {
            this.Ping();
            CellRef = rowRef;
            CellRefrences = SplitCellRefs(rowRef, this);
            this.Pong();
        }

        private static IEnumerable<CKCellRef> SplitCellRefs(CKRowCellRef rowRef, IDOMObject parent)
        {
            LH.Ping<CKRow>();
            var cellRefs = new List<CKCellRef>();
            for (int col_1 = 1; col_1 <= rowRef.ColumnIndex; col_1++)
            {
                cellRefs.Add(new CKCellRef(rowRef.RowIndex, col_1, rowRef.Table, parent));
            }
            LH.Pong<CKRow>();
            return cellRefs;
        }

        public CKRowCellRef CellRef { get; protected set; }
    }

    /// <summary>
    /// Represents a collection of <see cref="CKRow"/> objects in a Word table.
    /// This collection is part of the DOM hierarchy and implements <see cref="IDOMObject"/>.
    /// </summary>
    /// <remarks>
    /// Version: CK2.00.02.0003
    /// </remarks>
    public class CKRows : CKDOMObject, IEnumerable<CKRow>
    {
        private readonly Base1List<CKRow> _rows_1 = new Base1List<CKRow>();

        /// <summary>
        /// Initializes a new instance of the <see cref="CKRows"/> class.
        /// </summary>
        /// <param name="rows">The row collection to wrap.</param>
        /// <param name="parent">The owning parent DOM object.</param>
        /// <exception cref="ArgumentNullException">Thrown if <paramref name="rows"/> or <paramref name="parent"/> is null.</exception>
        public CKRows(IDOMObject parent)
        {
            this.Ping();
            //_rows_1 = rows ?? throw new ArgumentNullException(nameof(rows));
            Parent = parent ?? throw new ArgumentNullException(nameof(parent));
            this.Pong();
        }
        internal void Add(CKRow row)
        {
            this.Ping();
            _rows_1.Add(row);
            this.Pong();
        }

        /// <inheritdoc/>
        public override IDOMObject Parent { get; protected set; }

        /// <inheritdoc/>
        public override bool IsDirty
        {
            get => throw new NotImplementedException();//_isDirty || _rows_1.Any(r => r.IsDirty) || Parent.IsDirty;
            protected set => _isDirty = value;
        }
        bool _isDirty = false;
        /// <inheritdoc/>
        public override bool IsOrphan
        {
            get => throw new NotImplementedException();// _isOrphan || _rows_1.All(r => r.IsOrphan);
            protected set => _isOrphan = value;

        }
        bool _isOrphan = false;

        /// <summary>
        /// Gets the <see cref="CKRow"/> at the specified one-based index.
        /// </summary>
        /// <param name="index">The one-based row index.</param>
        /// <returns>The row at the specified index.</returns>
        /// <exception cref="ArgumentOutOfRangeException">Thrown if the index is invalid.</exception>
        public CKRow this[int index]
        {
            get
            {
                this.Ping(msg: $"Calling down to {typeof(Base1List<int>).Name}");

                var row = _rows_1[index];
                this.Pong();
                return row;
            }
        }

        /// <inheritdoc/>
        public IEnumerator<CKRow> GetEnumerator() => _rows_1.GetEnumerator();

        /// <inheritdoc/>
        IEnumerator IEnumerable.GetEnumerator() => GetEnumerator();

        /// <summary>
        /// Gets the number of rows in the collection.
        /// </summary>
        public int Count => _rows_1.Count;

        /// <inheritdoc/>
        public override string ToString() => $"CKRows [Count: {Count}]";
    }
}
