using System;
using System.Linq;
using Word = Microsoft.Office.Interop.Word;

namespace CoverageKiller2
{
    /// <summary>
    /// Provides methods for manipulating a Word table.
    /// </summary>
    public class CKWordTable
    {
        private readonly Word.Table _table;

        /// <summary>
        /// Initializes a new instance of the <see cref="CKWordTable"/> class.
        /// </summary>
        /// <param name="table">The Word table to wrap.</param>
        /// <exception cref="ArgumentNullException">Thrown when the provided table is null.</exception>
        public CKWordTable(Word.Table table)
        {
            _table = table ?? throw new ArgumentNullException(nameof(table), "Table cannot be null.");
        }

        /// <summary>
        /// Removes columns from the table based on the specified header texts.
        /// </summary>
        /// <param name="tabSeparatedHeaderTexts">A tab-separated string of header texts to remove.</param>
        public void RemoveColumnsByHeader(string tabSeparatedHeaderTexts)
        {
            // Split the tab-separated string into an array of headers to match
            var headersToRemove = tabSeparatedHeaderTexts.Split(new[] { '\t' }, StringSplitOptions.RemoveEmptyEntries);

            // Ensure to remove columns in reverse order to avoid shifting issues
            for (int i = _table.Columns.Count; i >= 1; i--)
            {
                string cellText = _table.Cell(1, i).Range.Text.Trim('\r', '\a'); // Get the header text in the first row

                // Check if the current cell text matches any of the headers to remove
                if (headersToRemove.Contains(cellText, StringComparer.OrdinalIgnoreCase))
                {
                    _table.Columns[i].Delete(); // Remove the column
                }
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
            if (rowIndex < 1 || rowIndex > _table.Rows.Count)
            {
                throw new ArgumentOutOfRangeException(nameof(rowIndex), "Row index is out of range.");
            }

            // Set the new value in the specified cell
            Word.Cell cell = _table.Cell(rowIndex, columnIndex);
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
            for (int i = 1; i <= _table.Columns.Count; i++)
            {
                string cellText = _table.Cell(1, i).Range.Text.Trim('\r', '\a'); // Get the header text in the first row

                // Compare ignoring case
                if (cellText.Equals(heading, StringComparison.OrdinalIgnoreCase))
                {
                    return i; // Return the index (1-based)
                }
            }

            return -1; // Return -1 if the heading is not found
        }
    }
}
