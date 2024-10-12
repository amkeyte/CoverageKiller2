using CoverageKiller2;
using System;
using Word = Microsoft.Office.Interop.Word;

public class TableFinder
{
    private Word.Application _application;
    private CKDocument _ckDoc;
    private string[] _headerTexts; // Changed to string array for header texts
    private Word.Table _currentTable;
    private int _currentTableIndex;

    public TableFinder(CKDocument ckDoc, string tabSeparatedHeaderTexts)
    {
        _application = ckDoc.WordApp;
        _ckDoc = ckDoc;
        _headerTexts = tabSeparatedHeaderTexts.Split(new[] { '\t' }, StringSplitOptions.RemoveEmptyEntries); // Split string into array
        _currentTableIndex = -1;
    }

    public bool TryFind(out Word.Table foundTable)
    {
        _currentTableIndex = -1; // Reset the current index
        foundTable = null; // Reset the found table

        for (int i = 1; i <= _ckDoc.Tables.Count; i++)
        {
            Word.Table table = _ckDoc.Tables[i];
            if (IsMatch(table))
            {
                foundTable = table;
                _currentTable = table;
                _currentTableIndex = i;
                return true; // Found the table
            }
        }

        return false; // No matching table found
    }

    public bool TryFindNext(out Word.Table foundTable)
    {
        if (_currentTableIndex == -1)
        {
            // Find the first matching table if none found yet
            return TryFind(out foundTable);
        }

        // Continue searching from the current index
        for (int i = _currentTableIndex + 1; i <= _ckDoc.Tables.Count; i++)
        {
            Word.Table table = _ckDoc.Tables[i];
            if (IsMatch(table))
            {
                foundTable = table;
                _currentTable = table;
                _currentTableIndex = i;
                return true; // Found the next matching table
            }
        }

        foundTable = null; // Reset the found table
        return false; // No more matching tables found
    }

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
            string cellText = firstRow.Cells[i + 1].Range.Text.Trim('\r', '\a'); // Remove end of cell characters
            if (!cellText.Equals(_headerTexts[i], StringComparison.OrdinalIgnoreCase))
            {
                return false; // Found a mismatch
            }
        }

        return true; // All header texts matched in sequence
    }
}
