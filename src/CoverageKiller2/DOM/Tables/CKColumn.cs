using CoverageKiller2.Logging;
using Serilog;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using Word = Microsoft.Office.Interop.Word;

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
            if (parent == null) throw new ArgumentNullException(nameof(parent));
            if (table == null) throw new ArgumentNullException(nameof(table));
            if (!table.Document.Equals(parent.Document)) throw new ArgumentException("table and parent must have the same document.");
            Index = colIndex;
            Table = table;
            Parent = parent;
            AccessMode = accessMode;
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

            CellRef = colRef;
            CellRefrences_1 = SplitCellRefs(colRef, this);
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

            return result;
        }
        /// <summary>
        /// Deletes the column if no merged cells exist, or falls back to SlowDelete.
        /// </summary>
        public void Delete()
        {
            var table = CellRef.Table;
            table.Columns.Delete(this);
        }



        public void SlowDeleteBak()
        {
            LH.Debug("Tracker[!sd]* Starting horizontal merge delete");

            var wordApp = Document.Application.WordApp;

            try
            {
                if (Index <= 1) throw new InvalidOperationException("Cannot merge left from column 1.");
                if (Count == 0) return;

                wordApp.ScreenUpdating = false;

                var progress = new LongOperationHelpers.ProgressLogger("CKColumn.MergeLeftDelete", Count);

                for (int i = 1; i <= Count; i++)
                {
                    var rightCell = this[i];

                    try
                    {
                        var row = rightCell.RowIndex;
                        var leftCell = rightCell.CellRef.Table.GetCellsFor(CellRef).FirstOrDefault();
                        if (leftCell != null) throw new CKDebugException("Cell not retrieved.");

                        LH.Debug($"[!sd] Merging Row {row}: Left ({Index - 1}) <- Right ({Index})");
                        leftCell.Merge(rightCell);
                    }
                    catch (COMException ex)
                    {
                        Log.Warning(ex, $"[!sd] Merge failed at row {i}");
                    }

                    progress.Report();

                    if (i % 100 == 0)
                    {
                        Log.Debug($"[!sd] Saving document after {i} row merges...");
                        LongOperationHelpers.TrySilentSave(Document, $"after {i} left merges");
                    }
                }

                progress.Finish();
                LongOperationHelpers.TrySilentSave(Document, "MergeLeftAndDeleteColumn: End of column merge operation.");
            }
            finally
            {
                wordApp.ScreenUpdating = true;
            }

            this.Clear();
            this.IsDirty = true;
            Log.Debug("Tracker [!sd]* Done tracking (Merge Left)");
        }

        public void SlowDelete()
        {
            LH.Debug("Tracker[!sd]* Starting slow delete");

            var wordApp = Document.Application.WordApp;

            try
            {
                if (Count == 0) return;

                wordApp.ScreenUpdating = false;

                // Pre-load COMCells once
                var comCells = new List<Word.Cell>();
                foreach (var cell in this.Cells)
                {
                    if (cell != null)
                    {
                        LH.Debug("Tracker[!sd] - looping");
                        comCells.Add(cell.COMCell);
                    }
                }

                Log.Debug($"[!sd][CKColumn.SlowDelete] Preparing to merge {comCells.Count} cells.");

                int totalMerged = 0;
                int batchSize = 10;

                var progress = new LongOperationHelpers.ProgressLogger("CKColumn.SlowDelete-Merge", comCells.Count);

                for (int i = 0; i < comCells.Count; i += batchSize)
                {
                    Word.Cell mergedCell = null;

                    int limit = Math.Min(batchSize, comCells.Count - i);
                    for (int j = 0; j < limit; j++)
                    {
                        var currentCell = comCells[i + j];
                        if (mergedCell == null)
                        {
                            LH.Debug("Tracker[!sd] - merge loop start");
                            mergedCell = currentCell;
                        }
                        else
                        {
                            mergedCell.Merge(currentCell);
                        }
                    }

                    if (mergedCell != null)
                    {
                        mergedCell.Delete();
                        Log.Debug($"[!sd] Batch deleted ({i + 1} to {i + limit})");
                    }

                    totalMerged += limit;
                    progress.Report();

                    if (totalMerged % 100 == 0)
                    {
                        Log.Debug($"[!sd] Saving document after {totalMerged} merged cells...");
                        LongOperationHelpers.TrySilentSave(Document, $"after {totalMerged} merged cells");
                    }
                }

                progress.Finish();
                LongOperationHelpers.TrySilentSave(Document, $"SlowDelete: End of column delete operation.");

                Log.Debug("[!sd] All batches processed.");
            }
            finally
            {
                wordApp.ScreenUpdating = true;
            }

            this.Clear();
            this.IsDirty = true;
            Log.Debug("Tracker [!sd]* Done tracking");
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
            Parent = parent ?? throw new ArgumentNullException(nameof(parent));
        }

        internal void Add(CKColumn column)
        {
            _columns_1.Add(column);
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
                var row = _columns_1[index];
                return row;
            }
        }

        public int Count => _columns_1.Count;

        public IEnumerator<CKColumn> GetEnumerator() => _columns_1.GetEnumerator();

        IEnumerator IEnumerable.GetEnumerator() => GetEnumerator();

        public override string ToString() => $"CKColumns [Count: {Count}]";

        /// <summary>
        /// Deletes the column at the specified one-based index.
        /// </summary>
        /// <param name="index">The one-based index of the column to delete.</param>
        public void Delete(int index)
        {

            if (index < 1 || index > Count)
                throw new ArgumentOutOfRangeException(nameof(index), "Index out of valid range.");

            var column = this[index];
            Delete(column);

        }

        /// <summary>
        /// Deletes the specified column from the table.
        /// </summary>
        /// <param name="column">The column to delete.</param>
        public void Delete(CKColumn column)
        {

            if (column == null)
                throw new ArgumentNullException(nameof(column));

            if (!_columns_1.Contains(column))
                throw new ArgumentException("Specified column does not exist in this collection.", nameof(column));

            var table = column.CellRef.Table;

            // Perform COM deletion here
            if (!table.HasMerge)
            {
                SafeCOM.Execute(table,
                    maxRetries: 1,
                    rethrow: false,
                    action: () =>
                    {
                        table.COMTable.Columns[column.Index].Delete();
                    });
            }
            else
            {
                //TODO more thought on this later.
                column.SlowDelete();

            }

            // Remove column from local cache
            _columns_1.Remove(column);

            IsDirty = true;

        }


        /// <summary>
        /// Deletes all the specified columns from the table.
        /// </summary>
        /// <param name="columns">The collection of columns to delete.</param>
        public void Delete(IEnumerable<CKColumn> columns)
        {

            if (columns == null) throw new ArgumentNullException(nameof(columns));

            var targets = columns.ToList();

            if (!targets.Any())
            {
                Log.Debug("No columns provided for deletion.");
                return;
            }

            //var targetHeaders = targets.Select(col =>
            //{
            //    try { return col[2].ScrunchedText; }
            //    catch { return "[Error getting text]"; }
            //});

            //LH.Debug($"Batch deleting columns: {targetHeaders.DumpString()}");

            for (int i = targets.Count - 1; i >= 0; i--)
            {
                Delete(targets[i]);
            }

            Log.Debug($"Batch deleted {targets.Count} columns. Remaining Count: {Count}");

            IsDirty = true;

        }
        /// <summary>
        /// Deletes all columns matching the given predicate.
        /// </summary>
        /// <param name="predicate">A function that evaluates each column.</param>
        public void Delete(Func<CKColumn, bool> predicate)
        {
            if (predicate == null) throw new ArgumentNullException(nameof(predicate));

            Delete(this.Where(predicate));
        }

    }
}
