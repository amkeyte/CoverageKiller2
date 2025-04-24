using CoverageKiller2.Logging;
using Serilog;
using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace CoverageKiller2.DOM.Tables
{


    /// <summary>
    /// Represents a one-based cell reference to a specific column in a Word table.
    /// </summary>
    public class CKColCellRef : CKCellRef, ICellRef<CKColumn>
    {
        //remember here that the underlying 
        public CKColCellRef(int colIndex, CKTable table, IDOMObject parent,
            TableAccessMode accessMode = TableAccessMode.IncludeAllCells)
            : base(1, colIndex, table, parent)//first cell as filler.
        {
            this.Ping();
            if (parent == null) throw new ArgumentNullException(nameof(parent));
            if (table == null) throw new ArgumentNullException(nameof(table));
            if (!table.Document.Equals(parent.Document)) throw new ArgumentException("table and parent must have the same document.");
            Index = colIndex;
            Table = table;
            Parent = parent;
            AccessMode = accessMode;
            this.Pong();
        }

        public int Index { get; }
        public CKTable Table { get; }
        public IDOMObject Parent { get; }
        public TableAccessMode AccessMode { get; }
    }

    /// <summary>
    /// Represents a column in a Word table.
    /// </summary>
    public class CKColumn : CKCells
    {
        public CKColumn(CKColCellRef colRef, IDOMObject parent)
            : base(parent)
        {
            this.Ping();

            CellRef = colRef;
            CellRefrences_1 = SplitCellRefs(colRef, this);
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

        public CKColCellRef CellRef { get; protected set; }
        public int Index => CellRef.Index;

        private IEnumerable<CKCellRef> SplitCellRefs(CKColCellRef colRef, IDOMObject parent)
        {
            this.Ping();
            //split out the ColRef into its individual cells.
            var cellRefs_1 = new Base1List<CKCellRef>();
            var rowCount = colRef.Table.GridRowCount;

            for (int row_1 = 1; row_1 <= rowCount; row_1++)
            {
                cellRefs_1.Add(new CKCellRef(row_1, colRef.ColumnIndex, colRef.Table, parent));
            }

            //filter for merged based on AccessMode.
            var result = cellRefs_1.Where(cr => cr.Table.FitsAccessMode(cr));
            IsDirty = true;

            return this.Pong(() => result, msg: result.ToString());
        }
        /// <summary>
        /// Deletes the column if no merged cells exist, or falls back to SlowDelete.
        /// </summary>
        public void Delete()
        {



            var topCell = CellsList_1[1];
            var table = topCell.Tables[1];

            Log.Debug($"Deleting column Index: {Index}");
            Log.Debug($"\tCollumn Document: {Document.FileName}::{Document.Content.Snapshot.FastHash}");
            Log.Debug($"\ttopCell Document: {topCell.Document.FileName}::{topCell.Document.Content.Snapshot.FastHash}");
            var comTblDoc = table.COMTable.Range.Document;
            Log.Debug($"\tCOMtable Document: {Path.GetFileName(comTblDoc.FullName)}::{new RangeSnapshot(comTblDoc.Content).FastHash}");
            var comRngDoc = table.COMRange.Document;
            Log.Debug($"\tCOMtable Document: {Path.GetFileName(comRngDoc.FullName)}::{new RangeSnapshot(comRngDoc.Content).FastHash}");



            if (!table.HasMerge)
            {

                table.COMTable.Columns[Index].Delete();
            }
            else
            {
                SlowDelete();
            }

            IsDirty = true;
            Log.Debug($"Deleted column: Index{Index}");
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
            var colIndex = CellRef.ColumnIndex;

            for (var i_1 = CellsList_1.Count; i_1 >= 1; i_1--)
            {
                var cell = CellsList_1[i_1];
                if (cell.ColumnIndex == colIndex)
                    cell.COMCell.Delete();
            }

            this.Pong();
        }
    }


    /// <summary>
    /// Represents a collection of <see cref="CKColumn"/> objects in a Word table.
    /// </summary>
    /// <remarks>
    /// Version: CK2.00.02.0003
    /// </remarks>
    public class CKColumns : CKDOMObject, IEnumerable<CKColumn>
    {
        private readonly Base1List<CKColumn> _columns_1 = new Base1List<CKColumn>();
        private bool _isDirty = false;
        private bool _isOrphan = false;

        public CKColumns(IDOMObject parent)
        {
            this.Ping();
            Parent = parent ?? throw new ArgumentNullException(nameof(parent));
            this.Pong();
        }

        internal void Add(CKColumn column)
        {
            this.Ping();
            _columns_1.Add(column);
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

        public int Count => _columns_1.Count;

        public IEnumerator<CKColumn> GetEnumerator() => _columns_1.GetEnumerator();

        IEnumerator IEnumerable.GetEnumerator() => GetEnumerator();

        public override string ToString() => $"CKColumns [Count: {Count}]";

        /// <summary>
        /// Deletes all columns matching the given predicate.
        /// </summary>
        /// <param name="predicate">A function that evaluates each column.</param>
        /// <returns>The list of columns that were deleted.</returns>
        public void Delete(Func<CKColumn, bool> predicate)
        {
            this.Ping();
            Log.Debug($"Deleting columns {Document.FileName}::{Document.Content.Snapshot.FastHash}::Before Count:{Count}");

            this.Where(predicate)
           .Reverse().ToList()
           .ForEach(col => col.Delete());

            Log.Debug($"Deleting columns {Document.FileName}::{Document.Content.Snapshot.FastHash}::After Count:{Count}");

            IsDirty = true;
            this.Pong();
        }
    }
}
