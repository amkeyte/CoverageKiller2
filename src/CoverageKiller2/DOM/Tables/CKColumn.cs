using CoverageKiller2.Logging;
using Serilog;
using System;
using System.Collections;
using System.Collections.Generic;
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

        /// <summary>
        /// Gets the CKCell by visual row index (1-based), matching the logical table row structure.
        /// </summary>
        public override CKCell this[int index]
        {
            get
            {
                //method updated to a search to keep indexer aligned with visual grid as expected. (issue 1)
                if (index < 1) throw new ArgumentOutOfRangeException(nameof(index));
                var cell = this.FirstOrDefault(c => c.RowIndex == index);
                if (cell == null)
                    throw new ArgumentOutOfRangeException($"No cell found at visual row {index} in column.");
                return cell;
            }
        }

        //get the column as a flat list if coordinate semantics are innapropriate
        public CKCells Cells => new CKCells(this, Parent);


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
            var table = CellRef.Table;



            Log.Debug($"[Issue1] Deleting column {Index} from " +
               $"{LH.GetTableTitle(table, "***Table")}" +
                   $".{Document.FileName}.{table.Snapshot}" +
                   $".Cell({CellsList_1[1]?.Snapshot})");



            if (!table.HasMerge)
            {

                table.COMTable.Columns[Index].Delete();
                Log.Debug($"[Issue1] Fast deleted column: Index{Index}");
            }
            else
            {
                SlowDelete();
            }

            IsDirty = true;
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

            LH.Debug($"[Issue1](slow) Deleting column {Index} from " +
                $"{LH.GetTableTitle(table, "***Table")}" +
                    $".{Document.FileName}.{table.Snapshot}" +
                    $".Cell({CellsList_1[1]?.Snapshot})");

            var colIndex = CellRef.ColumnIndex;

            for (var i_1 = CellsList_1.Count; i_1 >= 1; i_1--)
            {
                var cell = CellsList_1[i_1];
                if (cell.ColumnIndex == colIndex)
                {
                    cell.COMCell.Delete();
                    Log.Debug($"[Issue1] Deleted cell: Column {Index} Row {i_1}");

                }
            }

            Log.Debug($"[Issue1]Deleted column: Index{Index}");
            IsDirty = true;
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
            Log.Debug($"[Issue1] Batch Deleting columns {Document.FileName}::{Document.Content.Snapshot}::Before Count:{Count}");


            var x = this.Where(predicate).Select(col => col[2].ScrunchedText);
            LH.Debug($"[Issue1] The following column headers are slated for deletion:{x.DumpString()}");

            this.Where(predicate)
               .Reverse().ToList()
               .ForEach(col => col.Delete());

            Log.Debug($"[Issue1] Batch Deleted columns {Document.FileName}::{Document.Content.Snapshot}::After Count:{Count}");

            IsDirty = true;
            this.Pong();
        }
    }
}
