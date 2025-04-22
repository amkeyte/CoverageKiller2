using CoverageKiller2.Logging;
using System;
using System.Collections;
using System.Collections.Generic;

namespace CoverageKiller2.DOM.Tables
{
    public class CKColCellRef : CKCellRef, ICellRef<CKColumn>
    {
        public CKColCellRef(int colIndex, CKTable table, IDOMObject parent) :
            base(table.Rows.Count, colIndex, table, parent)
        {
            this.Ping();
            if (parent == null) throw new ArgumentNullException(nameof(parent));
            if (table == null) throw new ArgumentNullException(nameof(table));
            Index = colIndex;
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

    public class CKColumn : CKCells
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
        public CKColumn(Base1List<CKCell> cells_1, CKColCellRef cellRef, IDOMObject parent) : base(cells_1, cellRef.Parent)
        {
            Parent = parent;
            CellRef = cellRef;
        }
        public CKColumn(CKColCellRef colRef, IDOMObject parent) :
            base(parent)
        {
            this.Ping();
            CellRef = colRef;
            CellRefrences = SplitCellRefs(colRef, this);
            this.Pong();
        }

        private static IEnumerable<CKCellRef> SplitCellRefs(CKColCellRef colRef, IDOMObject parent)
        {
            LH.Ping<CKColumn>();
            var cellRefs = new List<CKCellRef>();
            for (int row_1 = 1; row_1 <= colRef.RowIndex; row_1++)
            {
                cellRefs.Add(new CKCellRef(row_1, colRef.ColumnIndex, colRef.Table, parent));
            }
            LH.Pong<CKColumn>();
            return cellRefs;
        }

        public void Delete()
        {
            this.PingPong("Deleting column!! (just kidding)");

        }

        public CKColCellRef CellRef { get; protected set; }
    }

    /// <summary>
    /// Represents a collection of <see cref="CKColumn"/> objects in a Word table.
    /// This collection is part of the DOM hierarchy and implements <see cref="IDOMObject"/>.
    /// </summary>
    /// <remarks>
    /// Version: CK2.00.02.0003
    /// </remarks>
    public class CKColumns : CKDOMObject, IEnumerable<CKColumn>
    {
        private readonly Base1List<CKColumn> _columns_1 = new Base1List<CKColumn>();


        public CKColumns(IDOMObject parent)
        {
            this.Ping();
            //_rows_1 = rows ?? throw new ArgumentNullException(nameof(rows));
            Parent = parent ?? throw new ArgumentNullException(nameof(parent));
            this.Pong();
        }
        internal void Add(CKColumn column)
        {
            this.Ping();
            _columns_1.Add(column);
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
        /// Gets the <see cref="CKColumn"/> at the specified one-based index.
        /// </summary>
        /// <param name="index">The one-based row index.</param>
        /// <returns>The row at the specified index.</returns>
        /// <exception cref="ArgumentOutOfRangeException">Thrown if the index is invalid.</exception>
        public CKColumn this[int index]
        {
            get
            {
                this.Ping(msg: $"Calling down to {typeof(Base1List<int>).Name}");

                var row = _columns_1[index];
                this.Pong();
                return row;
            }
        }

        /// <inheritdoc/>
        public IEnumerator<CKColumn> GetEnumerator() => _columns_1.GetEnumerator();

        /// <inheritdoc/>
        IEnumerator IEnumerable.GetEnumerator() => GetEnumerator();

        /// <summary>
        /// Gets the number of rows in the collection.
        /// </summary>
        public int Count => _columns_1.Count;

        /// <inheritdoc/>
        public override string ToString() => $"CKColumns [Count: {Count}]";
    }
}
