using System;
using Word = Microsoft.Office.Interop.Word;

namespace CoverageKiller2.Pipeline.WordHelpers
{
    /// <summary>
    /// Finds tables in a Word document based on specified header texts.
    /// </summary>
    public class TableFinder
    {
        private readonly Word.Application _application;
        private readonly CKDocument _ckDoc;
        private readonly string[] _headerTexts; // Array of header texts
        private Word.Table _currentTable;
        private int _currentTableIndex;

        /// <summary>
        /// Initializes a new instance of the <see cref="TableFinder"/> class.
        /// </summary>
        /// <param name="ckDoc">The CKDocument object that contains the Word tables.</param>
        /// <param name="tabSeparatedHeaderTexts">Tab-separated string of header texts to search for in the tables.</param>
        public TableFinder(CKDocument ckDoc, string tabSeparatedHeaderTexts)
        {
            _application = ckDoc.WordApp;
            _ckDoc = ckDoc;
            _headerTexts = tabSeparatedHeaderTexts.Split(new[] { '\t' }, StringSplitOptions.RemoveEmptyEntries); // Split headers by tab
            _currentTableIndex = -1; // No table found initially
        }

        /// <summary>
        /// Attempts to find the next table that matches the header texts. 
        /// If it's the first call, it searches for the first match.
        /// </summary>
        /// <param name="foundTable">Outputs the found table if a match is found.</param>
        /// <returns>True if a matching table is found, false otherwise.</returns>
        public bool TryFind(out Word.Table foundTable)
        {
            // If no table has been found yet, start searching from the beginning
            if (_currentTableIndex == -1)
            {
                _currentTableIndex = 0; // Reset to start
            }
            else
            {
                // Continue searching from the current index
                _currentTableIndex++;
            }

            // Search for a matching table
            for (int i = _currentTableIndex; i < _ckDoc.Tables.Count; i++)
            {
                Word.Table table = _ckDoc.Tables[i + 1]; // Note: Tables collection is 1-based
                if (IsMatch(table))
                {
                    foundTable = table;
                    _currentTable = table;
                    _currentTableIndex = i; // Update current index to the found table
                    return true; // Found a matching table
                }
            }

            foundTable = null; // No matching tables found
            return false;
        }

        /// <summary>
        /// Checks if the given table matches the specified header texts.
        /// </summary>
        /// <param name="table">The table to check.</param>
        /// <returns>True if the table matches the header texts, false otherwise.</returns>
        private bool IsMatch(Word.Table table)
        {
            if (table.Rows.Count < 1)
                return false;

            Word.Row firstRow = table.Rows[1]; // Get the first row (header row)

            // Check if the number of header texts matches the number of cells
            if (firstRow.Cells.Count < _headerTexts.Length)
                return false;

            // Ensure that all header texts match in sequence
            for (int i = 0; i < _headerTexts.Length; i++)
            {
                string cellText = firstRow.Cells[i + 1].Range.Text;//.Trim('\r', '\a'); // Remove end-of-cell characters
                if (!cellText.Equals(_headerTexts[i], StringComparison.OrdinalIgnoreCase))
                {
                    return false; // Found a mismatch
                }
            }

            return true; // All header texts matched in sequence
        }
    }
}
