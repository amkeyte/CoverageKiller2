using Serilog;
using System;
using System.Linq;
using System.Text.RegularExpressions;
using Word = Microsoft.Office.Interop.Word;

namespace CoverageKiller2
{
    /// <summary>
    /// Provides methods for manipulating a Word table.
    /// </summary>
    public class CKTable
    {
        internal static CKTable Create(CKTables parent, int tableIndex)
        {
            return new CKTable(parent, tableIndex);
        }

        /// <summary>
        /// do something when the table has been deleted??
        /// </summary>
        internal Word.Table COMObject { get; private set; }

        //the index might change if the table is altered, but pulling from parent index of will 
        // always return a current value.
        internal int _lastIndex;
        public int Index
        {
            get
            {
                Log.Debug(LH.TraceCaller(LH.PP.Enter, null,
                    nameof(CKTable), nameof(Index)));

                _lastIndex = Parent.ToList().IndexOf(this);

                Log.Debug(LH.TraceCaller(LH.PP.PropertyGet, null,
                    nameof(CKTable), nameof(Index),
                    $"{LH.ObjectPath(nameof(Parent), "ToList()", "IndexOf(this)")} --> ", _lastIndex));

                return _lastIndex;
            }
        }






        private CKTable(CKTables parent, int index)
        {
            _lastIndex = index;


            //here we store a reference to the com table itself in case
            // the document moves it in the index.
            // saving by index and repeatedly calling by that would resuly
            //in accessing the wrong table.
            // attempts to access this CKTable after it's deleted is an error.
            //Subsequent calls to a Tables[x] will return whatever table is indexed by x,
            //which is possibly not the same every time.
            Parent = parent;
            COMObject = Parent.COMObject[index];
        }

        //public bool ContainsMerged => Rows.ContainsMerged;
        public CKColumns Columns => CKColumns.Create(this);

        public CKRows Rows
        {
            get
            {
                Log.Debug(LH.TraceCaller(LH.PP.Enter, null,
                    nameof(CKTable), nameof(Rows)));

                var result = CKRows.Create(this);

                Log.Debug(LH.TraceCaller(LH.PP.PropertyGet, null,
                    nameof(CKTable), nameof(Rows),
                    $"{LH.ObjectPath(nameof(Rows))} --> ", result));

                return result;
            }
        }
        //=> CKRows.Create(this);

        private CKTables _parent;
        public CKTables Parent
        {
            get
            {
                Log.Debug(LH.TraceCaller(LH.PP.Enter, null,
                    nameof(CKTable), nameof(Parent)));

                var result = _parent;

                Log.Debug(LH.TraceCaller(LH.PP.PropertyGet, null,
                    nameof(CKTable), nameof(Parent),
                    $"{nameof(Parent)} --> ", result));

                return result;
            }
            private set
            {
                var setValue = value;
                Log.Debug(LH.TraceCaller(LH.PP.PropertySet, null,
                    nameof(CKTable), nameof(Parent),
                    $"{nameof(Parent)} --> ", setValue));

                _parent = value;

            }
        }

        /// <summary>
        /// Sets the value of a specified cell in the table.
        /// </summary>
        /// <param name="heading">The header text of the column to set the value for.</param>
        /// <param name="rowIndex">The row index of the cell to modify (1-based).</param>
        /// <param name="newValue">The new value to set in the cell.</param>
        /// <exception cref="ArgumentException">Thrown when the specified heading is not found in the table.</exception>
        /// <exception cref="ArgumentOutOfRangeException">Thrown when the row index is out of range.</exception>
        public void SetCell(string heading, int rowIndex, string newValue)
        {
            // Find the column index for the given heading
            int columnIndex = FindColumnIndexByHeading(heading);

            if (columnIndex == -1)
            {
                throw new ArgumentException($"Header '{heading}' not found in the table.");
            }

            // Check if the specified row index is valid
            if (rowIndex < 1 || rowIndex > COMObject.Rows.Count)
            {
                throw new ArgumentOutOfRangeException(nameof(rowIndex), "Row index is out of range.");
            }

            // Set the new value in the specified cell
            Word.Cell cell = COMObject.Cell(rowIndex, columnIndex);
            cell.Range.Text = newValue; // Replace the cell's text
        }

        /// <summary>
        /// Finds the index of the column with the specified heading.
        /// </summary>
        /// <param name="heading">The heading to find in the table.</param>
        /// <returns>The 1-based index of the column if found; otherwise, -1.</returns>
        private int FindColumnIndexByHeading(string heading)
        {
            // Loop through the columns in the first row to find the heading
            for (int i = 1; i <= COMObject.Columns.Count; i++)
            {
                string cellText = COMObject.Cell(1, i).Range.Text.Trim('\r', '\a'); // Get the header text in the first row

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

            COMObject.PreferredWidthType = Word.WdPreferredWidthType.wdPreferredWidthPercent;
            COMObject.PreferredWidth = 100f;
            //Log.Debug("Result {Type}, {Width}", COMObject.PreferredWidthType, COMObject.PreferredWidth);
        }

        /// <summary>
        /// Removes the table from the Word document.
        /// </summary>
        /// <exception cref="InvalidOperationException">Thrown if the table does not exist.</exception>
        public void Delete()
        {
            Log.Debug(LH.TraceCaller(LH.PP.Enter, null,
                nameof(CKTable), nameof(Delete),
                nameof(Index), _lastIndex));

            COMObject.Delete();
        }

        public bool RowMatches(int oneBasedRowIndex, string target)
        {
            if (oneBasedRowIndex <= 0 || oneBasedRowIndex > COMObject.Rows.Count)
                throw new ArgumentOutOfRangeException(nameof(oneBasedRowIndex), "Invalid row index.");

            // Combine all cell values in the row into one string
            var rowValues = string.Concat(
                COMObject.Rows[oneBasedRowIndex].Cells
                .Cast<Word.Cell>()
                .Select(cell => cell.Range.Text));

            string normalizedRowValues = NormalizeMatchString(rowValues);

            string normalizedTarget = NormalizeMatchString(target);

            //Log.Debug("{func}: row => {row}\n\ttarget => {target}\n\trowvalues => {rowVals}",
            //    nameof(RowMatches), oneBasedRowIndex, normalizedTarget, normalizedRowValues);

            // Compare the normalized strings
            return normalizedRowValues == normalizedTarget;
        }

        private string NormalizeMatchString(string input)
        {
            return Regex.Replace(input, @"[\x07\s]+", string.Empty);
        }


        public CKCell Cell(int row, int column)
        {

            return CKCell.Create(this, row, column);// new CKCell(COMObject.Cell(row, column));
        }

        //shameless hack

        public void AddAndMergeFirstRow(string text = "")
        {
            // Step 1: Add a new row at the top (first position)
            var newRow = Rows.Add(Rows[1]);

            // Step 2: Merge the cells in the new first row across all columns
            int numberOfColumnsToMerge = Columns.Count; // Get the total number of columns
            var firstCellInRow = Cell(1, 1); // First cell in the new first row
            var lastCellInRowToMerge = Cell(1, numberOfColumnsToMerge); // Last cell in the new first row

            // Merge the cells in the new first row from column 1 to the last column
            firstCellInRow.Merge(lastCellInRowToMerge);
            Cell(1, 1).Text = text;
        }

    }
}
