
using System;
using System.Linq;
using Word = Microsoft.Office.Interop.Word;

public class TableFixer
{
    private Word.Table _table;

    public TableFixer(Word.Table table)
    {
        _table = table ?? throw new ArgumentNullException(nameof(table), "Table cannot be null.");
    }

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