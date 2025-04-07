using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using Word = Microsoft.Office.Interop.Word;

namespace CoverageKiller2.DOM.Tables
{
    public interface ICellRef<out T> where T : IDOMObject
    {
        CKTable Table { get; }
        IEnumerable<int> CellIndexes { get; }
        IDOMObject Parent { get; }
    }
    /// <summary>
    /// Represents a reference to a cell or group of cells within a Word table.
    /// </summary>
    /// <remarks>
    /// Version: CK2.00.00.0000
    /// </remarks>
    public class CKCellRef : ICellRef<CKCell>
    {
        /// <inheritdoc/>
        public CKTable Table { get; }

        /// <inheritdoc/>
        public IEnumerable<int> CellIndexes { get; }

        /// <inheritdoc/>
        public IDOMObject Parent { get; }

        /// <summary>
        /// Gets the one-based Word row index of the referenced cell.
        /// </summary>
        public int WordRow { get; }

        /// <summary>
        /// Gets the one-based Word column index of the referenced cell.
        /// </summary>
        public int WordCol { get; }

        /// <summary>
        /// Initializes a new instance of the <see cref="CKCellRef"/> class.
        /// </summary>
        /// <param name="wordCell">The Word cell to reference.</param>
        /// <param name="parent">
        /// The parent DOM object, typically the owning <see cref="CKTable"/> or <see cref="CKCells"/> collection.
        /// </param>
        /// <exception cref="ArgumentNullException">Thrown if <paramref name="wordCell"/> or <paramref name="parent"/> is null.</exception>
        public CKCellRef(Word.Cell wordCell, IDOMObject parent)
        {
            if (wordCell == null) throw new ArgumentNullException(nameof(wordCell));
            if (parent == null) throw new ArgumentNullException(nameof(parent));

            Table = CKTable.FromRange(wordCell.Range, parent);
            CellIndexes = new List<int>() { Table.IndexOf(wordCell) };
            WordRow = wordCell.Row.Index;
            WordCol = wordCell.Column.Index;
            Parent = parent;
        }
    }

    /// <summary>
    /// Represents a single cell in a Word table, with DOM wrappers and location metadata.
    /// </summary>
    /// <remarks>
    /// Version: CK2.00.00.0000
    /// </remarks>
    public class CKCell : CKRange
    {
        /// <summary>
        /// Gets the underlying Word.Cell COM object.
        /// </summary>
        public Word.Cell COMCell { get; }

        /// <summary>
        /// Gets the CKTable to which this cell belongs.
        /// </summary>
        public CKTable Table { get; }

        /// <summary>
        /// Gets the one-based row index of the cell in the Word table.
        /// </summary>
        public int WordRow { get; }

        /// <summary>
        /// Gets the one-based column index of the cell in the Word table.
        /// </summary>
        public int WordColumn { get; }

        /// <summary>
        /// Gets the <see cref="CKCellRef"/> that describes this cell.
        /// </summary>
        public CKCellRef CellRef { get; }

        /// <summary>
        /// Initializes a new instance of the <see cref="CKCell"/> class.
        /// </summary>
        /// <param name="table">The parent CKTable that owns this cell.</param>
        /// <param name="parent">The logical parent DOM object, usually another table or range collection.</param>
        /// <param name="wdCell">The underlying Word.Cell to wrap.</param>
        /// <param name="wordRow">One-based row index of the cell.</param>
        /// <param name="wordColumn">One-based column index of the cell.</param>
        /// <exception cref="ArgumentNullException">Thrown if any parameter is null.</exception>
        public CKCell(CKTable table, IDOMObject parent, Word.Cell wdCell, int wordRow, int wordColumn)
            : base(wdCell?.Range ?? throw new ArgumentNullException(nameof(wdCell)), parent)
        {
            Table = table ?? throw new ArgumentNullException(nameof(table));
            COMCell = wdCell;
            WordRow = wordRow;
            WordColumn = wordColumn;
            CellRef = new CKCellRef(COMCell, parent);
        }

        /// <summary>
        /// Gets or sets the background color of the cell.
        /// </summary>
        public Word.WdColor BackgroundColor
        {
            get => COMCell.Shading.BackgroundPatternColor;
            set => COMCell.Shading.BackgroundPatternColor = value;
        }

        /// <summary>
        /// Gets or sets the foreground color of the cell.
        /// </summary>
        public Word.WdColor ForegroundColor
        {
            get => COMCell.Shading.ForegroundPatternColor;
            set => COMCell.Shading.ForegroundPatternColor = value;
        }
    }
    /// <summary>
    /// Represents a collection of <see cref="CKCell"/> instances derived from a <see cref="CKCellRef"/>.
    /// </summary>
    /// <remarks>
    /// Version: CK2.00.00.0000
    /// </remarks>
    public class CKCells : IEnumerable<CKCell>, IDOMObject
    {
        /// <summary>
        /// The owning table of the cell collection.
        /// </summary>
        public CKTable Table { get; protected set; }

        /// <summary>
        /// The original reference used to construct this collection.
        /// </summary>
        public CKCellRef CellRef { get; protected set; }

        /// <inheritdoc/>
        public IDOMObject Parent => Table;

        private readonly List<CKCell> _cells = new List<CKCell>();

        /// <summary>
        /// Protected constructor for subclassing or FromRef instantiation.
        /// </summary>
        protected CKCells() { }

        /// <summary>
        /// Builds the internal list of CKCell objects based on the supplied cell reference.
        /// </summary>
        /// <returns>Enumerable list of constructed CKCell objects.</returns>
        /// <exception cref="InvalidOperationException">Thrown if the CellRef or its index list is null.</exception>
        protected virtual IEnumerable<CKCell> BuildCells()
        {
            if (CellRef == null || CellRef.CellIndexes == null)
                throw new InvalidOperationException("CellRef or CellIndexes is null.");

            foreach (var i in CellRef.CellIndexes)
            {
                var gcr = Table.Converters.GetGridCellRef(i);
                var resolvedRef = Table.Converters.GetCellRef(gcr, CellRef.Parent);
                _cells.Add(Table.Cell(resolvedRef));
            }

            return _cells;
        }

        /// <summary>
        /// Constructs a new <see cref="CKCells"/> instance from a <see cref="CKTable"/> and a <see cref="CKCellRef"/>.
        /// </summary>
        /// <param name="table">The parent table.</param>
        /// <param name="cellRef">The reference to resolve into cell(s).</param>
        /// <returns>A new <see cref="CKCells"/> collection.</returns>
        /// <exception cref="ArgumentNullException">Thrown if either parameter is null.</exception>
        public static CKCells FromRef(CKTable table, CKCellRef cellRef)
        {
            if (table == null) throw new ArgumentNullException(nameof(table));
            if (cellRef == null) throw new ArgumentNullException(nameof(cellRef));

            var instance = new CKCells
            {
                Table = table,
                CellRef = cellRef
            };

            instance._cells.AddRange(instance.BuildCells());
            return instance;
        }

        /// <summary>
        /// Gets the number of cells in the collection.
        /// </summary>
        public int Count => _cells.Count;

        /// <inheritdoc/>
        public CKDocument Document => Table.Document;

        /// <inheritdoc/>
        public CKApplication Application => Parent.Application;

        /// <inheritdoc/>
        public bool IsDirty => Table.IsDirty || _cells.Any(c => c.IsDirty);

        /// <inheritdoc/>
        public bool IsOrphan => Document.IsOrphan;

        /// <summary>
        /// Gets the <see cref="CKCell"/> at the specified one-based index.
        /// </summary>
        /// <param name="index">The one-based index (1..Count).</param>
        /// <returns>The corresponding <see cref="CKCell"/> instance.</returns>
        /// <exception cref="ArgumentOutOfRangeException">If index is out of bounds.</exception>
        public CKCell this[int index]
        {
            get
            {
                if (index < 1 || index > Count)
                    throw new ArgumentOutOfRangeException(nameof(index));
                return _cells[index - 1];
            }
        }

        /// <inheritdoc/>
        public IEnumerator<CKCell> GetEnumerator() => _cells.GetEnumerator();

        /// <inheritdoc/>
        IEnumerator IEnumerable.GetEnumerator() => GetEnumerator();
    }

}