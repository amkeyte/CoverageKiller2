using CoverageKiller2.Logging;
using System;
using System.Collections;
using System.Collections.Generic;
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


        private CKTableGrid Grid => CKTableGrid.GetInstance(this, COMTable);
        private CKCellRefConverterService _converterService;

        /// <summary>
        /// Initializes a new instance of the <see cref="CKTable"/> class.
        /// </summary>
        /// <param name="table">The Word table to wrap.</param>
        /// <param name="parent">The logical parent DOM object.</param>
        public CKTable(Word.Table table, IDOMObject parent) : base(table.Range, parent)
        {
            LH.Ping(GetType());
            COMTable = table ?? throw new ArgumentNullException(nameof(table));
            _converterService = new CKCellRefConverterService(this);
            LH.Pong(GetType());
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
        public CKRows Rows => throw new NotImplementedException();

        /// <summary>
        /// Gets the columns of the table.
        /// </summary>
        public IEnumerable<CKColumn> Columns => throw new NotImplementedException();

        /// <summary>
        /// Gets the conversion helper service for this table.
        /// </summary>
        public CKCellRefConverterService Converters => _converterService;

        /// <summary>
        /// Retrieves the CKCell at the given reference.
        /// </summary>
        public CKCell Cell(CKCellRef cellRef)
        {
            var gridCellRef = Converters.GetGridCellRef(cellRef);
            //calling to Grid is required because it know if cells have been moved in the table.
            var gridCell = Grid.GetMasterCells(gridCellRef).FirstOrDefault()
                ?? throw new ArgumentException($"{nameof(cellRef)} did not fetch a master GridCell");


            return new CKCell(gridCell.COMCell, cellRef);
        }


        /// <summary>
        /// Returns the one-based index of a Word.Cell within the table.
        /// </summary>
        public int IndexOf(Word.Cell wordCell)
        {
            LH.Ping(GetType());
            var index = Cells.IndexOf(wordCell);
            LH.Pong(GetType());
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
            return new CKCell(gridCell.COMCell, cellRef);
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
            LH.Ping(GetType());
            COMTables = collection;
            LH.Pong(GetType());
        }

        private List<CKTable> TablesList
        {
            get
            {
                if (_cachedTables == null || IsDirty)
                {
                    _cachedTables = new List<CKTable>();
                    for (int i = 1; i <= COMTables.Count; i++)
                    {
                        _cachedTables.Add(new CKTable(COMTables[i], this));
                    }
                    IsDirty = false;
                }
                return _cachedTables;
            }
        }
        private List<CKTable> _cachedTables;

        /// <inheritdoc/>
        public override int Count => COMTables.Count;

        /// <inheritdoc/>
        private bool _isCheckingDirty = false;

        /// <inheritdoc/>
        public override bool IsDirty
        {
            get
            {
                if (_isDirty || _isCheckingDirty)
                    return _isDirty;

                _isCheckingDirty = true;
                try
                {
                    _isDirty = _cachedTables?.Any(t => t.IsDirty) == true || Parent.IsDirty;
                }
                finally
                {
                    _isCheckingDirty = false;
                }

                return _isDirty;
            }
            protected set => _isDirty = value;
        }
        /// <inheritdoc/>
        public override bool IsOrphan => throw new NotImplementedException();

        private Word.Tables COMTables { get; }

        /// <summary>
        /// Gets the <see cref="CKTable"/> at the specified one-based index.
        /// </summary>
        public CKTable this[int index]
        {
            get
            {
                if (index < 1 || index > Count)
                    throw new ArgumentOutOfRangeException(nameof(index), "Index must be between 1 and the number of tables.");
                return TablesList[index - 1];
            }
        }

        /// <inheritdoc/>
        public override int IndexOf(object obj)
        {
            if (obj is CKTable table)
            {
                return TablesList.IndexOf(table);
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
            LH.Ping(GetType());
            if (cell == null) throw new ArgumentNullException(nameof(cell));

            var ckTable = TablesList.FirstOrDefault(t => t.Contains(cell))
                ?? throw new ArgumentOutOfRangeException(nameof(cell), "Cell is not contained in any known table.");
            LH.Pong(GetType());
            return ckTable;

        }

        /// <inheritdoc/>
        public override string ToString() => $"CKTables [Count: {Count}]";

        /// <inheritdoc/>
        public IEnumerator<CKTable> GetEnumerator() => TablesList.GetEnumerator();

        /// <inheritdoc/>
        IEnumerator IEnumerable.GetEnumerator() => GetEnumerator();
    }
}
