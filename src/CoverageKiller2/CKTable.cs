using CoverageKiller2.Logging;
using System;
using System.Collections.Generic;
using System.Linq;
using Word = Microsoft.Office.Interop.Word;

namespace CoverageKiller2
{
    /// <summary>
    /// Provides methods for manipulating a Word table.
    /// </summary>
    public class CKTable : CKRange
    {
        public Tracer Tracer { get; } = new Tracer(typeof(CKTable));

        /// <summary>
        /// Remove from external references. Will be hidden.
        /// </summary>
        internal Word.Table COMTable { get; private set; }

        internal CKTable(Word.Table table) : base(table.Range)
        {
            Tracer.Enabled = false;
            COMTable = table;
            Grid = new CKTableGrid(this);
        }

        private CKTableGrid Grid { get; set; }

        public CKRows Rows => Grid.Rows;
        //public CKTables Parent
        //{
        //    get
        //    {
        //        var result = Tracer.Trace(_parent);
        //        Tracer.Log("Property Returned", new DataPoints(nameof(Parent)));
        //        return result;

        //    }
        //    private set
        //    {
        //        _parent = value;

        //    }
        //}

        /// <summary>
        /// Sets the value of a specified cell in the table.
        /// </summary>
        /// <param name="heading">The header text of the column to set the value for.</param>
        /// <param name="rowIndex">The row index of the cell to modify (1-based).</param>
        /// <param name="newValue">The new value to set in the cell.</param>
        /// <exception cref="ArgumentException">Thrown when the specified heading is not found in the table.</exception>
        /// <exception cref="ArgumentOutOfRangeException">Thrown when the row index is out of range.</exception>
        //public void SetCell(string heading, int rowIndex, string newValue)
        //{
        //    // Find the column index for the given heading
        //    int columnIndex = FindColumnIndexByHeading(heading);

        //    if (columnIndex == -1)
        //    {
        //        throw new ArgumentException($"Header '{heading}' not found in the table.");
        //    }

        //    // Check if the specified row index is valid
        //    if (rowIndex < 1 || rowIndex > COMTable.Rows.Count)
        //    {
        //        throw new ArgumentOutOfRangeException(nameof(rowIndex), "Row index is out of range.");
        //    }

        //    // Set the new value in the specified cell
        //    Word.Cell cell = COMTable.Cell(rowIndex, columnIndex);
        //    cell.Range.Text = newValue; // Replace the cell's text
        //}

        /// <summary>
        /// Finds the index of the column with the specified heading.
        /// </summary>
        /// <param name="heading">The heading to find in the table.</param>
        /// <returns>The 1-based index of the column if found; otherwise, -1.</returns>
        private int FindColumnIndexByHeading(string heading)
        {
            // Loop through the columns in the first row to find the heading
            for (int i = 1; i <= COMTable.Columns.Count; i++)
            {
                string cellText = COMTable.Cell(1, i).Range.Text.Trim('\r', '\a'); // Get the header text in the first row

                // Compare ignoring case
                if (cellText.Equals(heading, StringComparison.OrdinalIgnoreCase))
                {
                    return i; // Return the index (1-based)
                }
            }

            return -1; // Return -1 if the heading is not found
        }

        /// <summary>
        /// Creates a table in the active Word document and sets its width to the full page width.
        /// </summary>
        /// <remarks>
        /// This method retrieves the active document, inserts a table with a specified number of rows and columns,
        /// fills the table with data, and adjusts its width to occupy 100% of the page width.
        /// </remarks>
        public void MakeFullPage()
        {
            //Log.Debug("Setting Table width");

            COMTable.PreferredWidthType = Word.WdPreferredWidthType.wdPreferredWidthPercent;
            COMTable.PreferredWidth = 100f;
            //Log.Debug("Result {Type}, {Width}", COMObject.PreferredWidthType, COMObject.PreferredWidth);
        }

        /// <summary>
        /// Removes the table from the Word document.
        /// </summary>
        /// <exception cref="InvalidOperationException">Thrown if the table does not exist.</exception>
        public void Delete()
        {
            //Tracer.Log("Deleting Table", new DataPoints(nameof(Index)));

            COMTable.Delete();
        }

        public bool RowMatches(int oneBasedRowIndex, string target)
        {
            throw new NotImplementedException("Down for maintenance.");
            //if (oneBasedRowIndex <= 0 || oneBasedRowIndex > COMTable.Rows.Count)
            //    throw new ArgumentOutOfRangeException(nameof(oneBasedRowIndex), "Invalid row index.");

            //// Combine all cell values in the row into one string
            //var rowValues = string.Concat(
            //    COMTable.Rows[oneBasedRowIndex].Cells
            //    .Cast<Word.Cell>()
            //    .Select(cell => cell.Range.Text));

            //string normalizedRowValues = NormalizeMatchString(rowValues);

            //string normalizedTarget = NormalizeMatchString(target);
            //Tracer.Log("Table search match strings", new DataPoints()
            //    .Add(nameof(normalizedRowValues), normalizedRowValues)
            //    .Add(nameof(normalizedTarget), normalizedTarget));
            ////Log.Debug("{func}: row => {row}\n\ttarget => {target}\n\trowvalues => {rowVals}",
            ////    nameof(RowMatches), oneBasedRowIndex, normalizedTarget, normalizedRowValues);

            //// Compare the normalized strings
            //return normalizedRowValues == normalizedTarget;
        }

        private string NormalizeMatchString(string input)
        {
            throw new NotImplementedException("Down for maintenance.");
            //return Regex.Replace(input, @"[\x07\s]+", string.Empty);
        }


        public CKCell Cell(int row, int column)
        {
            throw new NotImplementedException("Down for maintenance.");
        }

        //shameless hack

        public void AddAndMergeFirstRow(string text = "")
        {
            throw new NotImplementedException("Down for maintenance.");
            //// Step 1: Add a new row at the top (first position)
            //var newRow = Rows.Add(Rows[1]);

            //// Step 2: Merge the cells in the new first row across all columns
            //int numberOfColumnsToMerge = Columns.Count; // Get the total number of columns
            //var firstCellInRow = Cell(1, 1); // First cell in the new first row
            //var lastCellInRowToMerge = Cell(1, numberOfColumnsToMerge); // Last cell in the new first row

            //// Merge the cells in the new first row from column 1 to the last column
            //firstCellInRow.Merge(lastCellInRowToMerge);
            //Cell(1, 1).Text = text;
        }


        public void Refresh()
        {
            var changes = CKTableGrid.Refresh(this);

            foreach (var x in CKCellReferences)
            {
                AdjustCellReference(x, changes);
            }
        }

        private void AdjustCellReference(CKCellReference cellRef, IEnumerable<GridChange> changes)
        {
            // Process row changes first: sort by RowIndex ascending.
            var rowChanges = changes
                .Where(c => c.ChangeType == GridChangeType.RowInserted || c.ChangeType == GridChangeType.RowDeleted)
                .OrderBy(c => c.RowIndex)
                .ToList();

            foreach (var change in rowChanges)
            {
                switch (change.ChangeType)
                {
                    case GridChangeType.RowInserted:
                        // If the insertion occurred before the reference's starting row, shift both Y1 and Y2 downward.
                        if (change.RowIndex < cellRef.Y1)
                        {
                            cellRef.Y1++;
                            cellRef.Y2++;
                        }
                        // If the insertion is within the referenced region, adjust Y2.
                        else if (change.RowIndex <= cellRef.Y2)
                        {
                            cellRef.Y2++;
                        }
                        break;
                    case GridChangeType.RowDeleted:
                        // If the deletion occurred before the reference's starting row, shift both Y1 and Y2 upward.
                        if (change.RowIndex < cellRef.Y1)
                        {
                            cellRef.Y1 = Math.Max(cellRef.Y1 - 1, 1);
                            cellRef.Y2 = Math.Max(cellRef.Y2 - 1, cellRef.Y1);
                        }
                        // If the deletion is within the referenced region, adjust Y2.
                        else if (change.RowIndex <= cellRef.Y2)
                        {
                            cellRef.Y2 = Math.Max(cellRef.Y2 - 1, cellRef.Y1);
                        }
                        break;
                }
            }

            // Process column changes: sort by ColumnIndex ascending.
            var colChanges = changes
                .Where(c => c.ChangeType == GridChangeType.ColumnInserted || c.ChangeType == GridChangeType.ColumnDeleted)
                .OrderBy(c => c.ColumnIndex)
                .ToList();

            foreach (var change in colChanges)
            {
                switch (change.ChangeType)
                {
                    case GridChangeType.ColumnInserted:
                        if (change.ColumnIndex < cellRef.X1)
                        {
                            cellRef.X1++;
                            cellRef.X2++;
                        }
                        else if (change.ColumnIndex <= cellRef.X2)
                        {
                            cellRef.X2++;
                        }
                        break;
                    case GridChangeType.ColumnDeleted:
                        if (change.ColumnIndex < cellRef.X1)
                        {
                            cellRef.X1 = Math.Max(cellRef.X1 - 1, 1);
                            cellRef.X2 = Math.Max(cellRef.X2 - 1, cellRef.X1);
                        }
                        else if (change.ColumnIndex <= cellRef.X2)
                        {
                            cellRef.X2 = Math.Max(cellRef.X2 - 1, cellRef.X1);
                        }
                        break;
                }
            }

            // Process cell modifications if needed.
            foreach (var change in changes.Where(c => c.ChangeType == GridChangeType.CellModified))
            {
                // Here you can handle individual cell modifications, such as merge/split changes,
                // if they affect the referenced region. For now, we leave this empty.
            }
        }

        private static Dictionary<CKTable, IEnumerable<CKCellReference>> _cellReferences
            = new Dictionary<CKTable, IEnumerable<CKCellReference>>();

        public List<CKCellReference> CKCellReferences => _cellReferences[this].ToList();
        internal IEnumerable<CKCell> CellItems(CKCellReference cellReference)
        {
            throw new NotImplementedException();
        }
    }
}
