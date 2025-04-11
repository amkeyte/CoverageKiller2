using CoverageKiller2.Logging;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using Word = Microsoft.Office.Interop.Word;

namespace CoverageKiller2.DOM.Tables
{
    public interface ICellRef<out T> where T : IDOMObject
    {
        //CKTable Table { get; }

        IDOMObject Parent { get; }
    }

    public class CellsRef : ICellRef<CKCells>, IEnumerable<CKCellRef>
    {
        public CellsRef(IEnumerable<CKCellRef> cellRefs, IDOMObject parent)
        {
            _cellRefs = cellRefs.ToList();
            Parent = parent;

        }
        private List<CKCellRef> _cellRefs;

        public IDOMObject Parent { get; }

        public IEnumerator<CKCellRef> GetEnumerator() => _cellRefs.GetEnumerator();


        IEnumerator IEnumerable.GetEnumerator() => GetEnumerator();

    }
    /// <summary>
    /// Represents a reference to a cell or group of cells within a Word table.
    /// </summary>
    /// <remarks>
    /// Version: CK2.00.01.0021
    /// </remarks>
    public class CKCellRef : ICellRef<CKCell>
    {
        /// <inheritdoc/>
        public IEnumerable<int> CellIndexes { get; }

        /// <inheritdoc/>
        public IDOMObject Parent { get; }

        /// <summary>
        /// Gets the one-based Word row index of the referenced cell.
        /// </summary>
        public int RowIndex { get; }

        /// <summary>
        /// Gets the one-based Word column index of the referenced cell.
        /// </summary>
        public int ColumnIndex { get; }

        /// <summary>
        /// Gets the snapshot of the original cell range, if captured.
        /// </summary>
        public RangeSnapshot Snapshot { get; }

        /// <summary>
        /// Initializes a new instance of the <see cref="CKCellRef"/> class with an optional snapshot.
        /// </summary>
        /// <param name="rowIndex">The one-based row index.</param>
        /// <param name="colIndex">The one-based column index.</param>
        /// <param name="snapshot">The snapshot of the original Word range, or null if not captured.</param>
        /// <param name="parent">The owning DOM object (table or collection).</param>
        public CKCellRef(int rowIndex, int colIndex, RangeSnapshot snapshot, IDOMObject parent)
        {
            if (parent == null) throw new ArgumentNullException(nameof(parent));

            RowIndex = rowIndex;
            ColumnIndex = colIndex;
            Snapshot = snapshot;
            Parent = parent;
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="CKCellRef"/> class without a snapshot.
        /// </summary>
        /// <param name="rowIndex">The one-based row index.</param>
        /// <param name="colIndex">The one-based column index.</param>
        /// <param name="parent">The owning DOM object (table or collection).</param>
        public CKCellRef(int rowIndex, int colIndex, IDOMObject parent)
            : this(rowIndex, colIndex, null, parent)
        {
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
        //public CKTable Table { get; }

        /// <summary>
        /// Gets the one-based row index of the cell in the Word table.
        /// </summary>
        public int RowIndex => CellRef.RowIndex;

        /// <summary>
        /// Gets the one-based column index of the cell in the Word table.
        /// </summary>
        public int ColumnIndex => CellRef.ColumnIndex;

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
        public CKCell(Word.Cell wdCell, CKCellRef cellRef)
            : base(wdCell?.Range, cellRef?.Parent)
        {
            //Table = new CKTable(wdCell.Tables[1], parent);
            //Table = table ?? throw new ArgumentNullException(nameof(table));
            COMCell = wdCell;
            CellRef = cellRef;
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
    public class CKCells : ACKRangeCollection, IEnumerable<CKCell>
    {
        /// <summary>
        /// The owning table of the cell collection.
        /// </summary>
        //public CKTable Table { get; protected set; }

        /// <summary>
        /// The original reference used to construct this collection.
        /// </summary>
        public CellsRef CellRef { get; protected set; }


        private List<CKCell> CellsList
        {
            get
            {
                LH.Ping(GetType());
                if (_cachedCells.Count == 0 || IsDirty)
                {
                    _cachedCells = new List<CKCell>();
                    for (int i = 1; i <= COMCells.Count; i++)
                    {
                        CKTable table = Document.Tables.ItemOf(COMCells[i]);

                        //var table = new CKTable(COMCells[i].Tables[1], Document);
                        var cellRef = new CKCellRef(
                            COMCells[i].RowIndex,
                            COMCells[i].ColumnIndex,
                            new RangeSnapshot(COMCells[i].Range),
                            this);

                        _cachedCells.Add(table.Cell(cellRef));
                    }
                    IsDirty = false;
                }
                LH.Pong(GetType());
                return _cachedCells;
            }
        }


        public CKCells(Word.Cells cells, CellsRef cellRef) : base(cellRef?.Parent)
        {
            COMCells = cells;
            CellRef = cellRef;
        }



        /// <inheritdoc/>

        public override int Count => COMCells.Count;
        bool _isDirty = false;
        private bool _isCheckingDirty = false;

        public override bool IsDirty
        {
            get
            {
                if (_isDirty || _isCheckingDirty)
                    return _isDirty;

                _isCheckingDirty = true;
                try
                {
                    _isDirty = _cachedCells.Any(c => c.IsDirty) || Parent.IsDirty;
                }
                finally
                {
                    _isCheckingDirty = false;
                }

                return _isDirty;
            }
            protected set => _isDirty = value;
        }

        public override bool IsOrphan => throw new NotImplementedException();

        public Word.Cells COMCells { get; private set; }

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

                return CellsList[index - 1];
            }
        }

        private List<CKCell> _cachedCells = new List<CKCell>();

        /// <inheritdoc/>
        public IEnumerator<CKCell> GetEnumerator() => CellsList.GetEnumerator();

        /// <inheritdoc/>
        IEnumerator IEnumerable.GetEnumerator() => GetEnumerator();

        public override int IndexOf(object obj)
        {
            LH.Ping(GetType());
            int index = -1;
            CKCell foundCell = default;
            if (obj is CKCell ckCell)
            {
                foundCell = (CKCell)obj;
            }
            else if (obj is Word.Cell wdCell)
            {
                var foundCells = CellsList.Where(c => RangeSnapshot.FastMatch(c.COMRange, wdCell.Range));
                foundCell = foundCells.FirstOrDefault();

            }

            index = CellsList.IndexOf(foundCell);
            LH.Pong(GetType());
            return index < 0 ? index : index + 1;
        }
    }

}