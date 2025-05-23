﻿using CoverageKiller2.Logging;
using Serilog;
using System;
using System.Collections.Generic;
using System.Linq;
using Word = Microsoft.Office.Interop.Word;

namespace CoverageKiller2.DOM.Tables
{

    public enum TableAccessMode
    {
        //Default,              
        IncludeAllCells,     // Uses merged cells as-is (current behavior)
        IncludeOnlyAnchorCells, // Includes only master cells (ignores merged duplicates)
        ExcludeAllMergedCells// Filters out any merged content
    }

    /// <summary>
    /// Provides methods for manipulating a Word table.
    /// </summary>
    /// <remarks>
    /// Version: CK2.00.00.0000
    /// </remarks>
    /// 
    public class CKTable : CKRange
    {
        public TableAccessMode AccessMode
        {
            get => Cache(ref _cachedAccessMode);
            set
            {
                _cachedAccessMode = value;
                IsDirty = true;
            }
        }
        public TableAccessMode _cachedAccessMode = TableAccessMode.IncludeAllCells;
        static CKTable()
        {
            LH.Ping<CKTable>(msg: $"Registering Caster for {nameof(CKTable)}");
            IDOMCaster.Register(input =>
            {
                LH.Ping<CKTable>(msg: $"Casting CKRange");

                if (!(input is CKRange inputRange))
                    throw new CKDebugException("input was not a range.");

                var doc = inputRange.Document;
                var tables = doc.Tables;

                var table = tables.FirstOrDefault(t => t.Equals(inputRange));

                if (table == null)
                    throw new CKDebugException($"A table was not matched in the document list for {doc.FileName} .");

                var result = new CKTable(table.COMTable, inputRange.Parent)
                    ?? throw new InvalidCastException("Could not convert to CKTable.");

                LH.Pong<CKTable>();
                return result;
            });
        }

        public CKTable(Word.Table table, IDOMObject parent) : base(table.Range, parent)
        {
            this.Ping();
            COMTable = table ?? throw new ArgumentNullException(nameof(table));
            _converterService = new CKCellRefConverterService(this);
            this.Pong();
        }

        public Word.Table COMTable { get; private set; }
        public CKCellRefConverterService Converters => _converterService;
        public int DocumentTableIndex => Document.Tables.IndexOf(this);
        internal bool HasMerge => Grid.HasMerge;
        internal bool FitsAccessMode(CKCellRef cellRef)
        {
            bool result = false;
            var gridCellRef = Converters.GetGridCellRef(cellRef);

            switch (AccessMode)
            {
                case TableAccessMode.IncludeAllCells:
                    result = true;
                    break;
                case TableAccessMode.IncludeOnlyAnchorCells:
                    result = !Grid.GetMergedCells(gridCellRef).Any();
                    break;
                case TableAccessMode.ExcludeAllMergedCells:
                    result = Grid.GetMasterCell(gridCellRef) != null;//hacked
                    break;
            }

            return this.Pong(() => result);
        }
        private CKTableGrid _grid;
        [Obsolete("Internal for testing only. Do not use in production code.")]
        internal CKTableGrid Grid => Cache(ref _grid, () => CKTableGrid.GetInstance(this, COMTable));

        private CKRows _rows_1;
        public CKRows Rows => this.PingPong(() => Cache(ref _rows_1, BuildRows));

        private CKColumns _cols_1;
        public CKColumns Columns => this.PingPong(() => Cache(ref _cols_1, BuildColumns));

        private readonly CKCellRefConverterService _converterService;

        /// <summary>
        /// Possibly slow
        /// </summary>
        /// <param name="cell"></param>
        /// <returns></returns>
        /// <exception cref="ArgumentNullException"></exception>
        public bool Contains(Word.Cell cell)
        {
            if (cell == null) throw new ArgumentNullException(nameof(cell));
            if (!this.Snapshot.SlowMatch(cell.Range.Tables[1].Range)) return true;
            try
            {
                var cellRef = new CKCellRef(cell.RowIndex, cell.ColumnIndex, new RangeSnapshot(cell.Range), this, this);
                var gridRef = Converters.GetGridCellRef(cellRef);
                return Grid.GetMasterCell(gridRef) != null; //hacked
            }
            catch
            {
                return false;
            }
        }

        public Word.Cell GetCellFor(CKCellRef cellRef)
        {
            //LH.Debug("Tracker[!sd]");
            var gridCellRef = Converters.GetGridCellRef(cellRef);
            var gridCell = Grid.GetMasterCell(gridCellRef)//hacked
                ?? throw new ArgumentException($"{nameof(cellRef)} did not fetch a master GridCell");

            int row = gridCell.GridRow;
            int col = gridCell.GridCol;
            //LH.Debug("Tracker[!sd]", "COMTable_get");
            //LH.Debug("Tracker[!sd]", "COMTable_get");
            if (row > COMTable.Rows.Count || col > COMTable.Columns.Count)
                throw new ArgumentOutOfRangeException($"Cell ({row}, {col}) does not exist in COM table [Rows: {COMTable.Rows.Count}, Cols: {COMTable.Columns.Count}].");
            //LH.Debug("Tracker[!sd]", "COMTable_get");
            //LH.Debug("Tracker[!sd]", "COMTable_get");
            if (gridCell.GridRow > COMTable.Rows.Count || gridCell.GridCol > COMTable.Columns.Count)
                throw new CKDebugException($"COM cell [{gridCell.GridRow},{gridCell.GridCol}] is outside table bounds [{COMTable.Rows.Count},{COMTable.Columns.Count}].");

            Log.Debug($"Requesting Cell({gridCell.GridRow}, {gridCell.GridCol})");
            //LH.Debug("Tracker[!sd]", "COMTable_get");
            var COMCell = COMTable.Cell(gridCell.GridRow, gridCell.GridCol); // this may break out of range.
            //Log.Debug($"Requesting COMTable cell returned cell[{gridCell.GridRow},{gridCell.GridCol})" +
            //    $" returned cell text '{COMCell.Range.Text.Scrunch()}");
            this.Pong();
            return COMCell;
        }


        public CKCells GetCellsFor(CKCellRef cellRef)
        {
            this.Ping(msg: $"Table [{DocumentTableIndex}]");
            var gridCellRef = Converters.GetGridCellRef(cellRef);
            var gridCells = Grid.GetMasterCell(gridCellRef);//hacked

            if (gridCells == null)// || !gridCells.An)
                throw new ArgumentException($"{nameof(cellRef)} did not fetch a master GridCell");

            //hacked for single cell use
            var result = new List<CKCell>();
            var COMCell = COMTable.Cell(gridCells.GridRow, gridCells.GridCol);
            result.Add(new CKCell(COMCell, cellRef));

            //foreach (var gridCell in gridCells)
            //{
            //    
            //    result.Add(new CKCell(COMCell, cellRef));
            //}
            return new CKCells(result, cellRef.Parent);
        }




        public int GridRowCount => Grid.RowCount;
        public int GridColCount => Grid.ColCount;


        [Obsolete]
        public void MakeFullPage()
        {
            var pageSetup = COMTable.Application.ActiveDocument.PageSetup;
            float usableWidth = pageSetup.PageWidth - pageSetup.LeftMargin - pageSetup.RightMargin;

            COMTable.PreferredWidthType = Word.WdPreferredWidthType.wdPreferredWidthPoints;
            COMTable.PreferredWidth = usableWidth;
            COMTable.Rows.Alignment = Word.WdRowAlignment.wdAlignRowLeft;
        }



        public CKCell Cell(int index)
        {
            var gridCellRef = Converters.GetGridCellRef(index);
            var gridCell = Grid.GetMasterCell(gridCellRef)//.FirstOrDefault()
                ?? throw new ArgumentException($"{nameof(index)} did not fetch a master GridCell");

            var cellRef = Converters.GetCellRef(gridCellRef, this);
            return new CKCell(cellRef);
            //var COMCell = COMTable.Cell(gridCell.GridRow, gridCell.GridCol);
            //return new CKCell(COMCell, cellRef);
        }

        internal void AutoFitBehavior(Word.WdAutoFitBehavior wdAutoFitContent) => throw new NotImplementedException();

        private CKRows BuildRows()
        {
            this.Ping();
            var rowCount = Grid.RowCount;
            var rows = new CKRows(this);
            for (var rowIndex = 1; rowIndex <= rowCount; rowIndex++)
            {
                var rowRef = new CKRowCellRef(rowIndex, this, rows, accessMode: AccessMode);
                rows.Add(new CKRow(rowRef, rows));
            }


            this.Pong();

            return _rows_1 = rows;
        }

        private CKColumns BuildColumns()
        {
            this.Ping();
            var colCount = Grid.ColCount;
            var cols = new CKColumns(this);

            for (var colIndex = 1; colIndex <= colCount; colIndex++)
            {
                var colRef = new CKColCellRef(colIndex, this, cols, accessMode: AccessMode);
                cols.Add(new CKColumn(colRef, cols));
            }

            this.Pong();
            return _cols_1 = cols;
        }

        /// <summary>
        /// Provides conversion services for cell reference and grid mapping.
        /// </summary>
        public class CKCellRefConverterService
        {
            public CKCellRefConverterService(CKTable table) => Table = table;
            public CKTable Table { get; private set; }
            public CKTableGrid Grid => Table.Grid;
        }
    }

}