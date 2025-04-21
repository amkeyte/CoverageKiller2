using CoverageKiller2.Logging;
using Microsoft.Office.Interop.Word;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using Word = Microsoft.Office.Interop.Word;

namespace CoverageKiller2.DOM.Tables
{

    /// <summary>
    /// Provides methods for manipulating a Word table.
    /// </summary>
    /// <remarks>
    /// Version: CK2.00.00.0000
    /// </remarks>
    public class CKTable : CKRange
    {
        static CKTable()
        {
            IDOMCaster.Register(input =>
            {
                LH.Ping<CKTable>($"Registering Caster for {nameof(CKTable)}");
                CKTable result = default;

                CKRange inputRange = input as CKRange ??
                    throw new CKDebugException("input was not a range.");

                var doc = inputRange.Document;
                var table = doc.Tables.Where(t => t == inputRange).FirstOrDefault() ??
                    throw new CKDebugException($"A table was not matched in the document list for {doc.FileName} .");

                //TODO need a COM safe way to produce a new table(range?) object with a different Parent
                result = new CKTable(table.COMTable, inputRange.Parent) ??
                    throw new InvalidCastException("Could not convert to CKTable.");

                //CKTables tables = default;
                //if (input.Parent is CKDocument doc)
                //{
                //    Log.Debug("Parent is CKDocument");
                //    tables = doc.Tables;
                //}
                //else if (input.Parent is CKTables ptables)
                //{
                //    Log.Debug("Parent is CKTables");
                //    tables = ptables;
                //}
                //else if (input.Parent is CKRange rng)
                //{
                //    Log.Debug("Parent is CKRange");
                //    tables = rng.Tables;
                //}
                //else
                //{
                //    Log.Warning($"Unrecognized input.Parent type: {input.Parent?.GetType().Name ?? "null"}");
                //}



                //var hashes = string.Join("\n", tables
                //        .Select(t => $"{t.Document.Tables.IndexOf(t)}:{t.Snapshot.FastHash}"));

                //LH.Checkpoint($"\nProspective Range fast hash: {inputRange.Snapshot.FastHash}" +
                //$"\nTable fast hashes ({tables.Document.Tables.Count}:\n"
                //    + hashes, typeof(CKTable));

                //// Try to locate by range comparison
                //result = tables.Where(t => t.Equals(inputRange)).FirstOrDefault();

                LH.Pong<CKTable>();
                return result;
            });
        }


        internal CKTableGrid Grid => CKTableGrid.GetInstance(this, COMTable);
        private CKCellRefConverterService _converterService;

        /// <summary>
        /// Initializes a new instance of the <see cref="CKTable"/> class.
        /// </summary>
        /// <param name="table">The Word table to wrap.</param>
        /// <param name="parent">The logical parent DOM object.</param>
        public CKTable(Word.Table table, IDOMObject parent) : base(table.Range, parent)
        {
            this.Ping();
            COMTable = table ?? throw new ArgumentNullException(nameof(table));
            _converterService = new CKCellRefConverterService(this);
            this.Pong();
        }
        /// <summary>
        /// Determines whether the specified <see cref="Word.Cell"/> exists within this table.
        /// </summary>
        /// <param name="cell">The Word cell to check.</param>
        /// <returns><c>true</c> if the cell is part of this table; otherwise, <c>false</c>.</returns>
        /// <remarks>
        /// Version: CK2.00.01.0003
        /// </remarks>
        public bool Contains(Word.Cell cell)
        {
            if (cell == null) throw new ArgumentNullException(nameof(cell));
            if (!RangeSnapshot.FastMatch(COMRange, cell.Range.Tables[1].Range)) return false;

            try
            {
                var cellRef = new CKCellRef(
                    cell.RowIndex,
                    cell.ColumnIndex,
                    new RangeSnapshot(cell.Range),
                    this,
                    this
                );

                var gridRef = Converters.GetGridCellRef(cellRef);

                return Grid.GetMasterCells(gridRef).Any();
            }
            catch
            {
                return false;
            }
        }

        /// <summary>
        /// Remove from external references. Will be hidden.
        /// </summary>
        public Word.Table COMTable { get; private set; }

        /// <summary>
        /// Gets the rows of the table.
        /// </summary>
        public CKRows Rows
        {
            get
            {
                this.Ping();
                if (_rows_1 == null)
                {

                    var rowCount = Grid.RowCount;
                    var colCount = Grid.ColCount;
                    LH.Checkpoint($"Grid.RowCount: {rowCount}; Grid.ColCount: {colCount}");
                    _rows_1 = new CKRows(this);

                    for (var rowIndex = 1; rowIndex <= rowCount; rowIndex++)
                    {
                        var row_ref = new CKRowCellRef(rowIndex, this, _rows_1);

                        _rows_1.Add(new CKRow(row_ref, _rows_1));
                    }
                    LH.Checkpoint($"_rows_1.Count: {_rows_1.Count}");
                }

                this.Pong();
                return _rows_1;
            }
        }

        //private CKCell GetCellFor(CKRowCellRef row_ref, int columnIndex)
        //{
        //    LH.Ping($"Table [{DocumentTableIndex}]", GetType());
        //    var gridCellRef = Converters.GetGridCellRef(row_ref);
        //    //calling to Grid is required because it (will) know if cells have been moved in the table.
        //    var gridCell = Grid.GetMasterCells(gridCellRef).FirstOrDefault();
        //    if (gridCell == null)
        //    {
        //        if (Debugger.IsAttached)
        //            Debugger.Break();
        //        else
        //            throw new ArgumentException($"{nameof(row_ref)}[{columnIndex}] did not fetch a master GridCell");

        //    }


        //    var COMCell = COMTable.Cell(gridCell.GridRow, gridCell.GridCol);
        //    LH.Pong(GetType());
        //    return new CKCell(COMCell, row_ref);
        //}

        private CKRows _rows_1;
        private CKColumns _cols_1;
        /// <summary>
        /// Gets the columns of the table.
        /// </summary>
        public CKColumns Columns
        {
            get
            {
                this.Ping();

                if (_cols_1 == null)
                {

                    var rowCount = Grid.RowCount;
                    var colCount = Grid.ColCount;
                    LH.Checkpoint($"Grid.RowCount: {rowCount}; Grid.ColCount: {colCount}");
                    _cols_1 = new CKColumns(this);

                    for (var colIndex = 1; colIndex <= colCount; colIndex++)
                    {
                        var colRef = new CKColCellRef(colIndex, this, _cols_1);

                        _cols_1.Add(new CKColumn(colRef, _cols_1));
                    }
                    LH.Checkpoint($"_cols_1.Count: {_cols_1.Count}");
                }

                this.Pong();
                return _cols_1;
            }
        }

        /// <summary>
        /// Gets the conversion helper service for this table.
        /// </summary>
        public CKCellRefConverterService Converters => _converterService;

        public int DocumentTableIndex => Document.Tables.IndexOf(this);
        /// <summary>
        /// Retrieves the Word.Cell at the given reference.
        /// </summary>
        public Word.Cell GetCellFor(CKCellRef cellRef)
        {
            this.Ping($"Table [{DocumentTableIndex}]");
            var gridCellRef = Converters.GetGridCellRef(cellRef);
            //calling to Grid is required because it know if cells have been moved in the table.
            var gridCell = Grid.GetMasterCells(gridCellRef).FirstOrDefault();
            if (gridCell == null)
            {
                if (Debugger.IsAttached)
                    Debugger.Break();
                else
                    throw new ArgumentException($"{nameof(cellRef)} did not fetch a master GridCell");

            }

            //should be the single point of loading the cell from COM for this chain.
            var COMCell = COMTable.Cell(gridCell.GridRow, gridCell.GridCol);
            this.Pong();
            return COMCell;

        }
        public CKCells GetCellsFor(CKCellRef cellRef)
        {
            this.Ping($"Table [{DocumentTableIndex}]");

            var gridCellRef = Converters.GetGridCellRef(cellRef);
            var gridCells_0 = Grid.GetMasterCells(gridCellRef);

            if (gridCells_0 == null || !gridCells_0.Any())
            {
                if (Debugger.IsAttached)
                    Debugger.Break();
                else
                    throw new ArgumentException($"{nameof(cellRef)} did not fetch a master GridCell");
            }

            var result_0 = new List<CKCell>();
            foreach (var gridCell in gridCells_0)
            {
                var COMCell = COMTable.Cell(gridCell.GridRow, gridCell.GridCol);
                result_0.Add(new CKCell(COMCell, cellRef));
            }
            return new CKCells(result_0, cellRef.Parent);
        }
        [Obsolete]//use CKRange text system
        public string DebugText
        {
            get
            {
                return RawText.Replace("\\", "\\\\");
            }
        }
        /// <summary>
        /// Parses a Word cell-debug dump into rows of strings.
        /// </summary>
        /// <param name="debugText">The raw debug dump string.</param>
        /// <returns>List of rows, each a list of cell texts.</returns>
        /// <remarks>
        /// Version: CK2.00.01.0021
        /// </remarks>
        [Obsolete]//fllagged as testing only.
        public Base1JaggedList<string> ParsedDebugText
        {
            get
            {
                var rows = new Base1JaggedList<string>();
                var currentRow = new Base1List<string>();

                // Split by "\r\a"
                var parts = DebugText.Split(new[] { "\r\a" }, StringSplitOptions.None);

                foreach (var part in parts)
                {
                    if (string.IsNullOrWhiteSpace(part))
                    {
                        if (currentRow.Count > 0)
                        {
                            rows.Add(currentRow);
                            currentRow = new Base1List<string>();
                        }
                    }
                    else
                    {
                        currentRow.Add(part.Trim());
                    }
                }

                if (currentRow.Count > 0)
                    rows.Add(currentRow);

                return rows;
            }
        }

        public WdPreferredWidthType PreferredWidthType { get; internal set; }
        public float PreferredWidth { get; internal set; }



        /// <summary>
        /// Resizes the table to span the full page width by adjusting its preferred width and alignment.
        /// </summary>
        /// <remarks>
        /// Version: CK2.00.01.0035
        /// </remarks>
        [Obsolete]//find a better way to do this without polluting the Table class.
        public void MakeFullPage()
        {
            var pageSetup = COMTable.Application.ActiveDocument.PageSetup;

            float pageWidth = pageSetup.PageWidth;
            float leftMargin = pageSetup.LeftMargin;
            float rightMargin = pageSetup.RightMargin;

            float usableWidth = pageWidth - leftMargin - rightMargin;

            COMTable.PreferredWidthType = Word.WdPreferredWidthType.wdPreferredWidthPoints;
            COMTable.PreferredWidth = usableWidth;

            COMTable.Rows.Alignment = Word.WdRowAlignment.wdAlignRowLeft; // Or wdAlignRowCenter if preferred
        }
        /// <summary>
        /// Returns the one-based index of a Word.Cell within the table.
        /// </summary>
        public int IndexOf(Word.Cell wordCell)
        {
            this.Ping("$$$");
            var index = Cells.IndexOf(wordCell);
            this.Pong();
            return index;
        }

        /// <summary>
        /// Retrieves a CKCell by one-based linear index.
        /// </summary>
        public CKCell Cell(int index)
        {
            var gridCellRef = Converters.GetGridCellRef(index);

            //calling to Grid is required because it know if cells have been moved in the table.
            var gridCell = Grid.GetMasterCells(gridCellRef).FirstOrDefault()
                ?? throw new ArgumentException($"{nameof(index)} did not fetch a master GridCell");

            var cellRef = Converters.GetCellRef(gridCellRef, this);
            var COMCell = COMTable.Cell(gridCell.RowSpan, gridCell.ColSpan);

            return new CKCell(COMCell, cellRef);
        }

        internal void AutoFitBehavior(WdAutoFitBehavior wdAutoFitContent)
        {
            throw new NotImplementedException();
        }

        /// <summary>
        /// Provides conversion services for cell reference and grid mapping.
        /// </summary>
        public class CKCellRefConverterService
        {
            /// <summary>
            /// Initializes a new instance for the given CKTable.
            /// </summary>
            public CKCellRefConverterService(CKTable table)
            {
                Table = table;
            }

            /// <summary>
            /// The owning CKTable.
            /// </summary>
            public CKTable Table { get; private set; }

            /// <summary>
            /// The associated CKTableGrid for the table.
            /// </summary>
            public CKTableGrid Grid => Table.Grid;
        }
    }
    /// <summary>
    /// Represents a collection of <see cref="CKTable"/> objects associated with a <see cref="CKRange"/>.
    /// </summary>
    /// <remarks>
    /// Version: CK2.00.01.0001
    /// </remarks>
    public class CKTables : ACKRangeCollection, IEnumerable<CKTable>
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="CKTables"/> class from the specified parent range.
        /// </summary>
        /// <param name="parent">The parent <see cref="CKRange"/> object that owns the tables.</param>
        public CKTables(Word.Tables collection, IDOMObject parent) : base(parent)
        {
            this.Ping();
            _comTables = collection;
            this.Pong();
        }
        private Base1List<CKTable> TablesList_1
        {
            get
            {
                this.Ping();
                if (!_cachedTables_1.Any() || IsDirty)
                {
                    _cachedTables_1.Clear();
                    for (int i = 1; i <= COMTables.Count; i++)
                    {

                        _cachedTables_1.Add(new CKTable(COMTables[i], this));
                    }
                    IsDirty = false;
                }
                this.Pong();
                return _cachedTables_1;
            }
        }
        private Base1List<CKTable> _cachedTables_1 = new Base1List<CKTable>();
        public override void Clear()
        {
            _cachedTables_1.Clear();
        }
        /// <inheritdoc/>
        public override int Count => TablesList_1.Count;

        protected override bool CheckDirtyFor()
        {
            this.PingPong();
            //TODO what goes here?
            return false;
        }

        /// <inheritdoc/>
        public override bool IsOrphan => throw new NotImplementedException();


        private Word.Tables _comTables;
        private Word.Tables COMTables => this.PingPong(() => _comTables);

        /// <summary>
        /// Gets the <see cref="CKTable"/> at the specified one-based index.
        /// </summary>
        public CKTable this[int index]
        {
            get
            {
                this.Ping("$$$");
                if (index < 1 || index > Count)
                    throw new ArgumentOutOfRangeException(nameof(index), "Index must be between 1 and the number of tables.");
                var result = TablesList_1[index];
                this.Pong();
                return result;
            }
        }

        /// <inheritdoc/>
        public override int IndexOf(object obj)
        {
            if (obj is CKTable table)
            {
                return (TablesList_1.IndexOf(table));

                //for (int tableIndex_1 = 1; tableIndex_1 <= TablesList_1.Count; tableIndex_1++)
                //{
                //    if ( TablesList_1[tableIndex_1].Equals(  table)) return tableIndex_1;
                //}
            }
            return -1;
        }
        /// <summary>
        /// Adds a new <see cref="CKTable"/> at the specified range with the given number of rows and columns.
        /// </summary>
        /// <param name="insertAt">The <see cref="CKRange"/> at which to insert the table.</param>
        /// <param name="numRows">The number of rows in the new table.</param>
        /// <param name="numColumns">The number of columns in the new table.</param>
        /// <returns>The newly created <see cref="CKTable"/> instance.</returns>
        /// <exception cref="ArgumentNullException">If <paramref name="insertAt"/> is null.</exception>
        /// <exception cref="ArgumentOutOfRangeException">If <paramref name="numRows"/> or <paramref name="numColumns"/> is less than 1.</exception>
        /// <remarks>
        /// Version: CK2.00.01.0006
        /// </remarks>
        public CKTable Add(CKRange insertAt, int numRows, int numColumns)
        {
            if (insertAt == null) throw new ArgumentNullException(nameof(insertAt));
            if (numRows < 1) throw new ArgumentOutOfRangeException(nameof(numRows));
            if (numColumns < 1) throw new ArgumentOutOfRangeException(nameof(numColumns));
            if (insertAt.Start != insertAt.End) throw new ArgumentException($"{nameof(insertAt)} must be collapsed.");

            var wordTable = COMTables.Add(insertAt.COMRange, numRows, numColumns);

            IsDirty = true;
            return new CKTable(wordTable, this);
        }
        /// <summary>
        /// Returns the <see cref="CKTable"/> that owns the given <see cref="Word.Cell"/>, if present.
        /// </summary>
        /// <param name="cell">A Word cell to search for.</param>
        /// <returns>The owning <see cref="CKTable"/>.</returns>
        /// <exception cref="ArgumentOutOfRangeException">Thrown if no owning table is found.</exception>
        /// <remarks>
        /// Version: CK2.00.01.0004
        /// </remarks>
        internal CKTable ItemOf(Word.Cell cell)
        {
            this.Ping();
            if (cell == null) throw new ArgumentNullException(nameof(cell));

            var ckTable = TablesList_1.FirstOrDefault(t => t.Contains(cell))
                ?? throw new ArgumentOutOfRangeException(nameof(cell), "Cell is not contained in any known table.");
            this.Pong();
            return ckTable;

        }

        /// <inheritdoc/>
        public override string ToString() => $"CKTables [Count: {Count}]";

        /// <inheritdoc/>
        public IEnumerator<CKTable> GetEnumerator() => TablesList_1.GetEnumerator();

        /// <inheritdoc/>
        IEnumerator IEnumerable.GetEnumerator() => GetEnumerator();
    }
}
