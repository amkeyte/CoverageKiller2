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
        private readonly Word.Table _table;
        private bool _tableDeleted = false;

        public Word.Table WordTable
        {
            get
            {
                if (_tableDeleted) throw new NullReferenceException("The table has been deleted");
                if (_table is null) throw new NullReferenceException("A reference to the table does not exist. The CKTable wrapper may be out of sync.");
                return _table;
            }
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="CKTable"/> class.
        /// </summary>
        /// <param name="table">The Word table to wrap.</param>
        /// <exception cref="ArgumentNullException">Thrown when the provided table is null.</exception>
        public CKTable(Word.Table table)
        {
            _table = table ?? throw new ArgumentNullException(nameof(table), "Table cannot be null.");
        }

        public bool ContainsMerged => Rows.ContainsMerged;
        public CKColumns Columns => new CKColumns(_table.Columns);

        public CKRows Rows => new CKRows(_table.Rows);

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
            if (rowIndex < 1 || rowIndex > WordTable.Rows.Count)
            {
                throw new ArgumentOutOfRangeException(nameof(rowIndex), "Row index is out of range.");
            }

            // Set the new value in the specified cell
            Word.Cell cell = WordTable.Cell(rowIndex, columnIndex);
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
            for (int i = 1; i <= WordTable.Columns.Count; i++)
            {
                string cellText = WordTable.Cell(1, i).Range.Text.Trim('\r', '\a'); // Get the header text in the first row

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
            Log.Debug("Setting Table width");

            WordTable.PreferredWidthType = Word.WdPreferredWidthType.wdPreferredWidthPercent;
            WordTable.PreferredWidth = 100f;
            Log.Debug("Result {Type}, {Width}", WordTable.PreferredWidthType, WordTable.PreferredWidth);
        }

        /// <summary>
        /// Removes the table from the Word document.
        /// </summary>
        /// <exception cref="InvalidOperationException">Thrown if the table does not exist.</exception>
        public void Delete()
        {
            // Remove the table from the document
            WordTable.Delete();
            _tableDeleted = true;
        }

        public bool RowMatches(int oneBasedRowIndex, string target)
        {
            if (oneBasedRowIndex <= 0 || oneBasedRowIndex > WordTable.Rows.Count)
                throw new ArgumentOutOfRangeException(nameof(oneBasedRowIndex), "Invalid row index.");

            // Combine all cell values in the row into one string
            var rowValues = string.Concat(
                WordTable.Rows[oneBasedRowIndex].Cells
                .Cast<Word.Cell>()
                .Select(cell => cell.Range.Text));

            string normalizedRowValues = NormalizeMatchString(rowValues);

            string normalizedTarget = NormalizeMatchString(target);

            Log.Debug("{func}: row => {row}\n\ttarget => {target}\n\trowvalues => {rowVals}",
                nameof(RowMatches), oneBasedRowIndex, normalizedTarget, normalizedRowValues);

            // Compare the normalized strings
            return normalizedRowValues == normalizedTarget;
        }

        private string NormalizeMatchString(string input)
        {
            return Regex.Replace(input, @"[\x07\s]+", string.Empty);
        }
    }
}
