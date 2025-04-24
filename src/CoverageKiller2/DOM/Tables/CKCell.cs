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
        CKTable Table { get; }

        IDOMObject Parent { get; }
    }

    //[Obsolete]
    //public class CellRefs : IEnumerable<CKCellRef>
    //{
    //    public CellRefs(IEnumerable<CKCellRef> cellRefs, IDOMObject parent)
    //    {
    //        _cellRefs_0 = cellRefs.ToList();
    //        Parent = parent;

    //    }
    //    private List<CKCellRef> _cellRefs_0;
    //    public int Count => _cellRefs_0.Count;
    //    public IDOMObject Parent { get; }

    //    public IEnumerator<CKCellRef> GetEnumerator() => _cellRefs_0.GetEnumerator();


    //    IEnumerator IEnumerable.GetEnumerator() => GetEnumerator();

    //}
    /// <summary>
    /// Represents a reference to a cell or group of cells within a Word table.
    /// </summary>
    /// <remarks>
    /// Version: CK2.00.01.0021
    /// </remarks>
    public class CKCellRef : ICellRef<CKCell>
    {

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
        public CKCellRef(int rowIndex, int colIndex, RangeSnapshot snapshot, CKTable table, IDOMObject parent)
        {
            this.Ping();
            if (parent == null) throw new ArgumentNullException(nameof(parent));
            if (table == null) throw new ArgumentNullException(nameof(table));

            RowIndex = rowIndex;
            ColumnIndex = colIndex;
            Snapshot = snapshot;
            Table = table;
            Parent = parent;
            this.Pong();
        }
        public CKTable Table { get; }



        public CKCellRef(int rowIndex, int colIndex, CKTable table, IDOMObject parent)
            : this(rowIndex, colIndex, null, table, parent)
        {
            this.Ping();


            if (rowIndex < 1 || rowIndex > table.GridRowCount) throw new ArgumentOutOfRangeException(nameof(rowIndex));
            if (colIndex < 1 || colIndex > table.GridColCount) throw new ArgumentOutOfRangeException(nameof(colIndex));


            if (parent == null) throw new ArgumentNullException(nameof(parent));
            if (table == null) throw new ArgumentNullException(nameof(table));
            if (!table.Document.Matches(table.Document)) throw new ArgumentException("table and parent must share a document.");
            this.Pong();
        }
    }

    /// <summary>
    /// Represents a single cell in a Word table, with DOM wrappers and location metadata.
    /// </summary>
    /// <remarks>
    /// Version: CK2.00.02.0001
    /// </remarks>
    public class CKCell : CKRange
    {
        /// <summary>
        /// Gets the underlying Word.Cell COM object, with caching support.
        /// </summary>
        [Obsolete("Make protected")]
        public Word.Cell COMCell
        {
            get
            {
                this.Ping();
                if (_COMCell == null || IsDirty)
                {
                    var table = CellRef.Table;
                    _COMCell = table.GetCellFor(CellRef);
                    if (COMRange is null) COMRange = _COMCell.Range;
                    //COMRange = _COMCell.Range;
                }
                this.Pong();
                return _COMCell;
            }
        }
        private Word.Cell _COMCell;

        /// <summary>
        /// Ensures the underlying Word.Cell object is current.
        /// </summary>
        /// <exception cref="CKDebugException"></exception>
        protected override void DoRefreshThings()
        {
            this.Ping();
            //checked if it's null to force COMCell to update, so that COMRange is valid.
            if (COMCell == null) throw new CKDebugException("COMCell cannot refresh.");
            //base.Refresh();
            this.Pong();
        }
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
            _COMCell = wdCell;
            CellRef = cellRef;
        }
        public CKCell(CKCellRef cellRef) : base(cellRef?.Parent)
        {
            //Table = new CKTable(wdCell.Tables[1], parent);
            //Table = table ?? throw new ArgumentNullException(nameof(table));
            CellRef = cellRef;
            IsDirty = true;
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

        public IEnumerable<CKCellRef> CellRefrences_1 { get; protected set; }

        /// <summary>
        /// Broken for Rows
        /// </summary>
        protected Base1List<CKCell> CellsList_1
        {
            get
            {
                this.Ping();
                if (_cachedCells_1.Count == 0 || IsDirty)
                {
                    _cachedCells_1 = new Base1List<CKCell>();
                    foreach (var cellRef in CellRefrences_1)
                    {
                        _cachedCells_1.Add(new CKCell(cellRef));
                    }
                    IsDirty = false;
                }
                this.Pong();
                return _cachedCells_1;
            }
        }

        public override void Clear()
        {
            _cachedCells_1?.Clear();
        }

        public CKCells(IEnumerable<CKCell> cells, IDOMObject parent) : base(parent)
        {
            _cachedCells_1 = new Base1List<CKCell>(cells);
            CellRefrences_1 = new Base1List<CKCellRef>(cells.Select(c => c.CellRef));
            if (cells.Any(c => !Document.Matches(parent.Document)))
                throw new ArgumentException("cells and parent must have matching documents.");

            IsDirty = false; //the cells were just dmped in!
        }
        public CKCells(IDOMObject parent) : base(parent)
        {
            CellRefrences_1 = new Base1List<CKCellRef>();
            _cachedCells_1 = new Base1List<CKCell>();

            IsDirty = true;
        }


        /// <inheritdoc/>

        //public override int Count => COMCells.Count;
        private bool _isCheckingDirty = false;

        protected override bool CheckDirtyFor()
        {
            //TODO: when is Cells dirty? checking each cell is heavy.
            return false;
        }

        public override bool IsOrphan => throw new NotImplementedException();

        public override int Count => CellsList_1.Count;

        //public Word.Cells COMCells { get; private set; }

        /// <summary>
        /// Gets the <see cref="CKCell"/> at the specified one-based index.
        /// </summary>
        /// <param name="index">The one-based index (1..Count).</param>
        /// <returns>The corresponding <see cref="CKCell"/> instance.</returns>
        /// <exception cref="ArgumentOutOfRangeException">If index is out of bounds.</exception>
        public virtual CKCell this[int index]
        {
            get
            {
                if (index < 1 || index > Count)
                    throw new ArgumentOutOfRangeException(nameof(index));
                return CellsList_1[index];
            }
        }

        private Base1List<CKCell> _cachedCells_1 = new Base1List<CKCell>();

        /// <inheritdoc/>
        public IEnumerator<CKCell> GetEnumerator() => CellsList_1.GetEnumerator();

        /// <inheritdoc/>
        IEnumerator IEnumerable.GetEnumerator() => GetEnumerator();

        public override int IndexOf(object obj)
        {
            this.Ping();
            int index = -1; //default return not found.
            CKCell foundCell = default;
            if (obj is CKCell ckCell)
            {
                foundCell = (CKCell)obj;
            }
            else if (obj is Word.Cell wdCell)
            {
                var foundCells_0 = CellsList_1.Where(c => c.Snapshot.SlowMatch(wdCell.Range));
                foundCell = foundCells_0.FirstOrDefault();
            }

            index = CellsList_1.IndexOf(foundCell);
            this.Pong();
            return index;
        }
    }

}