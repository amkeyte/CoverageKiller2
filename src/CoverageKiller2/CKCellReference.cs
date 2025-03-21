using System;
using System.Linq;
using Word = Microsoft.Office.Interop.Word;

namespace CoverageKiller2
{
    public class CKCellReference
    {
        public enum ReferenceTypes
        {
            Coordinate,
            Row,
            Column,
            Range
        }

        public ReferenceTypes RefType { get; }

        // Coordinates represent column (X) and row (Y) numbers.
        private int _x1;
        private int _y1;
        private int _x2;
        private int _y2;

        public int X1
        {
            get { if (Table.IsDirty) Table.Refresh(); return _x1; }
            internal set { _x1 = value; }
        }

        public int Y1
        {
            get { if (Table.IsDirty) Table.Refresh(); return _y1; }
            internal set { _y1 = value; }
        }

        public int X2
        {
            get { if (Table.IsDirty) Table.Refresh(); return _x2; }
            internal set { _x2 = value; }
        }

        public int Y2
        {
            get { if (Table.IsDirty) Table.Refresh(); return _y2; }
            internal set { _y2 = value; }
        }

        private int KeepLive(int val)
        {
            //refresh calls back to setter to update if needed.
            if (Table.IsDirty) Table.Refresh(); return val;
        }


        /// <summary>
        /// Constructs a cell reference from two coordinate pairs.
        /// Throws an exception if coordinates are reversed or outside of the table bounds.
        /// </summary>
        public CKCellReference(
            CKTable table,
            int x1, int y1,
            int x2, int y2,
            ReferenceTypes refType = ReferenceTypes.Range)
        {
            if (table == null)
                throw new ArgumentNullException(nameof(table));

            // Ensure coordinates are provided in increasing order.
            if (x1 > x2 || y1 > y2)
                throw new ArgumentException("Coordinates must be provided in increasing order: x1 <= x2 and y1 <= y2.");

            // Validate coordinates against table boundaries.
            if (x1 < 1 || y1 < 1 ||
                x2 > table.Columns.Count || y2 > table.Rows.Count)
            {
                throw new ArgumentOutOfRangeException("Coordinates are outside of the table bounds.");
            }

            this.Range = table;
            Table.CKCellReferences.Add(this);

            this.RefType = refType;
            this.X1 = x1;
            this.Y1 = y1;
            this.X2 = x2;
            this.Y2 = y2;
        }

        public CKTable Table
        {
            get
            {
                if (Range.IsDirty) throw new InvalidOperationException("probably Broke here.");
                return Range.Tables.FirstOrDefault();
            }
        }

        public CKRange Range { get; private set; }

        /// <summary>
        /// Constructs a cell reference for a single cell.
        /// </summary>
        public CKCellReference(
            CKTable table,
            int x1, int y1)
        {
            Range = table;
            Table.CKCellReferences.Add(this);

            RefType = ReferenceTypes.Coordinate;

            // For a single cell, start and stop coordinates are the same.
            this.X1 = this.X2 = x1;
            this.Y1 = this.Y2 = y1;
        }

        /// <summary>
        /// Constructs a cell reference that represents an entire row or column.
        /// For rows, index is the row number (all columns in that row).
        /// For columns, index is the column number (all rows in that column).
        /// </summary>
        public CKCellReference(
            CKTable table, int index,
            ReferenceTypes refType)
        {
            Range = table;
            Table.CKCellReferences.Add(this);

            if (refType != ReferenceTypes.Row && refType != ReferenceTypes.Column)
            {
                throw new ArgumentException("refType must be either Row or Column", nameof(refType));
            }
            RefType = refType;

            if (refType == ReferenceTypes.Row)
            {
                // Select the entire row.
                this.X1 = 1;                         // first column
                this.X2 = table.Columns.Count;       // last column in the table
                this.Y1 = this.Y2 = index;             // the row number is fixed
            }
            else // Column
            {
                // Select the entire column.
                this.Y1 = 1;                         // first row
                this.Y2 = table.Rows.Count;          // last row in the table
                this.X1 = this.X2 = index;             // the column number is fixed
            }
        }


        /// <summary>
        /// Creates a cell reference from an arbitrary CKRange.
        /// The range might be exactly a table, larger than a table, or just part of one.
        /// We use the first table within the range.
        /// Then we determine the bounding cell coordinates (as columns and rows) for all cells
        /// in that table that overlap with the CKRange.
        /// </summary>
        public CKCellReference(CKRange range)
        {
            Range = range;
            RefType = ReferenceTypes.Range;

            if (Table == null)
            {
                // Optionally throw or handle the case when no table exists.
                X1 = Y1 = X2 = Y2 = 0;
                return;
            }

            Table.CKCellReferences.Add(this);

            int minRow = int.MaxValue;
            int minCol = int.MaxValue;
            int maxRow = int.MinValue;
            int maxCol = int.MinValue;

            // Iterate over each cell in the table.
            // We assume Table.COMTable is a Word.Table.
            foreach (Word.Cell cell in Table.COMTable.Range.Cells)
            {
                Word.Range cellRange = cell.Range;

                // Check if the cell overlaps the CKRange.
                // This simple overlap test considers any partial overlap.
                if (cellRange.Start < Range.COMRange.End && cellRange.End > Range.COMRange.Start)
                {
                    // Get the row index directly from the cell's Row property.
                    int rowIndex = cell.Row.Index;
                    int colIndex = GetColumnIndex(cell);

                    minRow = Math.Min(minRow, rowIndex);
                    minCol = Math.Min(minCol, colIndex);
                    maxRow = Math.Max(maxRow, rowIndex);
                    maxCol = Math.Max(maxCol, colIndex);
                }
            }

            if (minRow == int.MaxValue || minCol == int.MaxValue)
            {
                // No overlapping cells were found.
                X1 = Y1 = X2 = Y2 = 0;
            }
            else
            {
                // Set the bounding coordinates.
                X1 = minCol;
                Y1 = minRow;
                X2 = maxCol;
                Y2 = maxRow;
            }
        }

        /// <summary>
        /// Helper method to compute the column index for a given cell.
        /// Word does not expose a direct ColumnIndex, so we iterate over the cells in the row.
        /// </summary>
        private int GetColumnIndex(Word.Cell cell)
        {
            Word.Row row = cell.Row;
            int index = 1;
            foreach (Word.Cell c in row.Cells)
            {
                if (c == cell)
                {
                    return index;
                }
                index++;
            }
            // Should never reach here if the cell is indeed in the row.
            return -1;
        }


    }
}
