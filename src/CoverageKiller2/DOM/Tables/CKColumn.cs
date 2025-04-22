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
            var topCell = CellsList_1[1];
            var table = topCell.Tables[1];
            if (!table.HasMerge)
            {
                table.Columns[Index].Delete();
            }
            {
                SlowDelete();
            }
        }

        /// <summary>
        /// Deletes all non-merged cells in this column using the CKTable grid layout.
        /// This is a fallback method when Word's native column deletion fails due to merged cells.
        /// </summary>
        /// <remarks>
        /// Version: CK2.00.03.0001
        /// </remarks>
        public void SlowDelete()
        {
            this.Ping();

            var table = CellRef.Table;
            var grid = table.Grid;
            var colIndex = CellRef.ColumnIndex;

            //remove in reverse order
            for (var i_1 = CellsList_1.Count; i_1 >= 1; i_1--)
            {
                var cell = CellsList_1[i_1];

                if (cell.ColumnIndex == colIndex)
                    cell.COMCell.Delete();
            }
            //TODO do something about the dirty grid.


            //for (int rowIndex = 1; rowIndex <= grid.RowCount; rowIndex++)
            //{
            //    var gridCell = grid[rowIndex, colIndex];
            //    if (gridCell == null) continue;

            //    // Only delete master cells to avoid removing shared merged cells multiple times
            //    if (gridCell.IsMasterCell && !gridCell.IsMerged)
            //    {
            //        try
            //        {
            //            var cellRef = new CKCellRef(rowIndex, colIndex, table, this);
            //            var cell = table.GetCellFor(cellRef);
            //            cell.Delete(); // Use the COM delete directly
            //        }
            //        catch (Exception ex)
            //        {
            //            Log.Warning($"Failed to delete cell at ({rowIndex},{colIndex}): {ex.Message}", ex);
            //        }
            //    }
            //}

            this.Pong();
        }

        public int Index => CellRef.Index;//keep an eye on this for concurrency
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
