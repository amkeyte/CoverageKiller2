using CoverageKiller2.Logging;
using Serilog;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;

namespace CoverageKiller2.DOM.Tables
{

    /// <summary>
    /// Represents a one-based cell reference to a specific row in a Word table.
    /// </summary>
    public class CKRowCellRef : CKCellRef, ICellRef<CKRow>
    {
        public CKRowCellRef(int rowIndex, CKTable table, IDOMObject parent,
            TableAccessMode accessMode = TableAccessMode.IncludeAllCells)
            : base(rowIndex, table.GridColCount, table, parent)
        {
            this.Ping();


            Index = rowIndex;
            Table = table;
            Parent = parent;
            AccessMode = accessMode;
            this.Pong();
        }

        public int Index { get; }
        public CKTable Table { get; }
        public IDOMObject Parent { get; }
        public TableAccessMode AccessMode { get; set; }
    }

    /// <summary>
    /// Represents a row in a Word table.
    /// </summary>
    public class CKRow : CKCells
    {
        //public CKRow(Base1List<CKCell> cells_1, CKRowCellRef cellRef, IDOMObject parent)
        //    : base(cells_1, cellRef.Parent)
        //{
        //    Parent = parent;
        //    RowRef = cellRef;
        //}

        public CKRow(CKRowCellRef rowRef, IDOMObject parent)
            : base(parent)
        {
            this.Ping();
            RowRef = rowRef;
            CellRefrences_1 = SplitCellRefs(rowRef, this);//maybe do this lazy
            this.Pong();
        }

        public override CKCell this[int index]
        {
            get
            {
                if (index < 1 || index > Count)
                    throw new ArgumentOutOfRangeException(nameof(index));
                return CellsList_1[index];
            }
        }

        public CKRowCellRef RowRef { get; protected set; }
        /// <summary>
        /// Deletes the row if no merged cells exist, or falls back to SlowDelete.
        /// </summary>
        /// <remarks>
        /// Version: CK2.00.03.0002
        /// </remarks>
        public void Delete()
        {
            var leftCell = CellsList_1[1];
            var table = leftCell.Tables[1];

            if (!table.HasMerge)
            {
                table.COMTable.Rows[Index].Delete();
            }
            else
            {
                SlowDelete();
            }

            IsDirty = true;
            Log.Debug("Deleted row Index");
        }
        public int Index => RowRef.Index;
        /// <summary>
        /// Deletes all non-merged cells in this row using the CKTable grid layout.
        /// This is a fallback method when Word's native row deletion fails due to merged cells.
        /// </summary>
        /// <remarks>
        /// Version: CK2.00.03.0003
        /// </remarks>
        public void SlowDelete()
        {
            this.Ping();

            var table = RowRef.Table;
            var rowIndex = RowRef.RowIndex;

            for (var i = CellsList_1.Count; i >= 1; i--)
            {
                var cell = CellsList_1[i];
                if (cell.RowIndex == rowIndex)
                    cell.COMCell.Delete();
            }

            this.Pong();
        }

        private IEnumerable<CKCellRef> SplitCellRefs(CKRowCellRef rowRef, IDOMObject parent)
        {
            this.Ping();
            var cellRefs = new Base1List<CKCellRef>();
            var colCount = rowRef.Table.GridColCount;

            for (int col = 1; col <= colCount; col++)
            {
                cellRefs.Add(new CKCellRef(rowRef.RowIndex, col, rowRef.Table, parent));
            }


            var result = cellRefs.Where(cr => cr.Table.FitsAccessMode(cr));


            return this.Pong(() => result, msg: result.ToString());

        }
    }

    /// <summary>
    /// Represents a collection of <see cref="CKRow"/> objects in a Word table.
    /// </summary>
    /// <remarks>
    /// Version: CK2.00.03.0001
    /// </remarks>
    public class CKRows : CKDOMObject, IEnumerable<CKRow>
    {
        private readonly Base1List<CKRow> _rows_1 = new Base1List<CKRow>();
        private bool _isDirty = false;
        private bool _isOrphan = false;
        internal string DumpList => _rows_1.Dump();
        public CKRows(IDOMObject parent)
        {
            this.Ping();
            Parent = parent ?? throw new ArgumentNullException(nameof(parent));
            this.Pong();
        }

        internal void Add(CKRow row)
        {
            this.Ping();
            _rows_1.Add(row);
            this.Pong();
        }

        public override IDOMObject Parent { get; protected set; }

        public override bool IsDirty
        {
            get => throw new NotImplementedException();
            protected set => _isDirty = value;
        }

        public override bool IsOrphan
        {
            get => throw new NotImplementedException();
            protected set => _isOrphan = value;
        }

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

        public int Count => _rows_1.Count;

        public IEnumerator<CKRow> GetEnumerator() => _rows_1.GetEnumerator();

        IEnumerator IEnumerable.GetEnumerator() => GetEnumerator();

        public override string ToString() => $"CKRows [Count: {Count}]";
    }
}
