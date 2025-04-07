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
            var wordCell = COMTable.Cell(cellRef.WordRow, cellRef.WordCol);
            return new CKCell(this, cellRef.Parent, wordCell, gridCellRef.Y1 + 1, gridCellRef.X1 + 1);
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
            var wordCell = COMTable.Range.Cells[index];
            return new CKCell(this, this, wordCell, gridCellRef.Y1 + 1, gridCellRef.X1 + 1);
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
    /// Version: CK2.00.00.0000
    /// </remarks>
    public class CKTables : ACKRangeCollection, IEnumerable<CKTable>
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="CKTables"/> class from the specified parent range.
        /// </summary>
        /// <param name="parent">The parent <see cref="CKRange"/> object that owns the tables.</param>
        public CKTables(CKRange parent) : base(parent) { }

        /// <summary>
        /// Gets the number of tables in the associated range.
        /// </summary>
        public override int Count => COMTables.Count;

        /// <summary>
        /// Gets the underlying Word.Tables COM object from the parent range.
        /// </summary>
        public Word.Tables COMTables => Parent.COMRange.Tables;

        /// <summary>
        /// Gets the <see cref="CKTable"/> at the specified one-based index.
        /// </summary>
        /// <param name="index">The one-based index of the table to retrieve.</param>
        /// <returns>The <see cref="CKTable"/> at the specified index.</returns>
        /// <exception cref="ArgumentOutOfRangeException">
        /// Thrown when the index is less than 1 or greater than the number of tables.
        /// </exception>
        public CKTable this[int index]
        {
            get
            {
                if (index < 1 || index > Count)
                {
                    throw new ArgumentOutOfRangeException(nameof(index), "Index must be between 1 and the number of tables.");
                }
                return new CKTable(COMTables[index], this);
            }
        }

        /// <summary>
        /// Returns a string that represents the current <see cref="CKTables"/> instance.
        /// </summary>
        /// <returns>A string containing the count of tables.</returns>
        public override string ToString()
        {
            return $"CKTables [Count: {Count}]";
        }

        /// <summary>
        /// Gets whether this collection or any contained table is dirty.
        /// </summary>
        public override bool IsDirty => _isDirty || this.Any(x => x.IsDirty);

        public override bool IsOrphan => throw new NotImplementedException();

        /// <summary>
        /// Returns an enumerator that iterates through the <see cref="CKTable"/> objects in the collection.
        /// </summary>
        /// <returns>An enumerator for the collection of <see cref="CKTable"/> objects.</returns>
        public IEnumerator<CKTable> GetEnumerator()
        {
            for (int i = 1; i <= Count; i++)
            {
                yield return this[i];
            }
        }

        /// <summary>
        /// Returns an enumerator that iterates through the collection.
        /// </summary>
        /// <returns>An enumerator for the collection.</returns>
        IEnumerator IEnumerable.GetEnumerator() => GetEnumerator();
    }

}
