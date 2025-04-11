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
        /// <summary>
        /// Creates a CKTable from a Word range, using the provided parent to resolve document context.
        /// </summary>
        /// <param name="wordRange">The Word.Range to extract the table from.</param>
        /// <param name="parent">The logical parent DOM object.</param>
        /// <returns>A CKTable instance for the contained table.</returns>
        [Obsolete("Requires exposed COM object")]
        public static CKTable FromRange(Word.Range wordRange, IDOMObject parent)
        {
            if (wordRange == null) throw new ArgumentNullException(nameof(wordRange));
            if (wordRange.Tables.Count == 0)
                throw new ArgumentException($"{nameof(wordRange)} does not contain a table.");

            var doc = parent.Document;
            var foundTable = doc.Tables
                .Where(t => t.COMRange.Contains(wordRange))
                .FirstOrDefault();

            return foundTable ?? new CKTable(wordRange.Tables[1], parent);
        }

        private CKTableGrid Grid => CKTableGrid.GetInstance(this, COMTable);
        private CKCellRefConverterService _converterService;

        /// <summary>
        /// Initializes a new instance of the <see cref="CKTable"/> class.
        /// </summary>
        /// <param name="table">The Word table to wrap.</param>
        /// <param name="parent">The logical parent DOM object.</param>
        public CKTable(Word.Table table, IDOMObject parent) : base(table.Range, parent)
        {
            COMTable = table ?? throw new ArgumentNullException(nameof(table));
            _converterService = new CKCellRefConverterService(this);
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
        /// Returns one-based indexes of cells in the given Word.Cells collection.
        /// </summary>
        public IEnumerable<int> IndexesOf(Word.Cells wordCells)
        {
            var tableCellList = COMTable.Range.Cells.ToList();
            return wordCells.AsEnumerable()
                .Select(c => IndexOf(c, tableCellList));
        }

        /// <summary>
        /// Returns the one-based index of a Word.Cell within the table.
        /// </summary>
        public int IndexOf(Word.Cell wordCell, List<Word.Cell> tableCells = null)
        {
            var tableCellList = tableCells ?? COMTable.Range.Cells.ToList();
            return tableCellList.FindIndex(c => c.Range.COMEquals(wordCell.Range)) + 1;
        }

        /// <summary>
        /// Returns one-based indexes of cells in a CKCells collection.
        /// </summary>
        public IEnumerable<int> IndexesOf(CKCells cells)
        {
            IEnumerable<GridCell> masterCells = Grid.GetMasterCells();
            IEnumerable<GridCell> findCells = cells
                .Select(c => Converters.GetGridCellRef(c.CellRef))
                .Select(c => Grid.GetMasterCells(c).First());

            return findCells.Select(c => IndexOf(c, masterCells));
        }

        private int IndexOf(GridCell c, IEnumerable<GridCell> masterCells = null)
        {
            var masterList = new List<GridCell>(masterCells ?? Grid.GetMasterCells());
            return masterList.IndexOf(c);
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

            var cellRef = Converters.GetCellRef(gridCellRef);
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
            COMTables = collection;
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
        public override bool IsDirty { get; protected set; } = false;

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
            if (cell == null) throw new ArgumentNullException(nameof(cell));

            return TablesList.FirstOrDefault(t => t.Contains(cell))
                ?? throw new ArgumentOutOfRangeException(nameof(cell), "Cell is not contained in any known table.");
        }

        /// <inheritdoc/>
        public override string ToString() => $"CKTables [Count: {Count}]";

        /// <inheritdoc/>
        public IEnumerator<CKTable> GetEnumerator() => TablesList.GetEnumerator();

        /// <inheritdoc/>
        IEnumerator IEnumerable.GetEnumerator() => GetEnumerator();
    }
}
