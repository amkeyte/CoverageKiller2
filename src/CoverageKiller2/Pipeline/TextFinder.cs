using CoverageKiller2;
using System;
using Word = Microsoft.Office.Interop.Word;

public class TextFinder
{
    private readonly CKDocument _ckDoc; // Wrapper for Word.Document
    private Word.Range _currentRange;
    private Word.Range _searchWithinRange;
    private readonly string _searchText;
    public bool TextFound { get; private set; } = false;

    // Constructor accepts CKDocument and the search text
    public TextFinder(CKDocument ckDoc, string searchText) : this(ckDoc, searchText, null) { }

    public TextFinder(CKDocument ckDoc, string searchText, Word.Range searchWithinRange)
    {
        _ckDoc = ckDoc;
        _searchWithinRange = searchWithinRange ?? _ckDoc.Content; // Ensure valid range
        _currentRange = _searchWithinRange;
        _searchText = searchText;

    }

    // Try to find the first occurrence of the search text
    public bool TryFind(out Word.Range foundRange)
    {
        TextFound = false;
        _currentRange = _searchWithinRange; // Reset range to the search domain
        bool found = _currentRange.Find.Execute(FindText: _searchText, MatchWildcards: true);

        if (found)
        {
            foundRange = _currentRange;
            TextFound = true;
            return true;
        }
        else
        {
            foundRange = null;
            return false;
        }
    }

    public bool TryFindNext(out Word.Range foundRange, bool wrap = false)
    {
        TextFound = false;
        int originalStart = _currentRange.Start; // Save the original starting point
        _currentRange.Start = _currentRange.End; // Move range start to after the last found occurrence

        // Try to find the next occurrence
        bool found = _currentRange.Find.Execute(FindText: _searchText, MatchWildcards: true);

        // If not found and wrapping is enabled, start from the beginning of the document
        if (!found && wrap)
        {
            _currentRange = _searchWithinRange; // Reset to the search range
            found = _currentRange.Find.Execute(FindText: _searchText, MatchWildcards: true);

            // Ensure we don't match the same initial occurrence (avoid infinite loop)
            if (found && _currentRange.Start == originalStart)
            {
                found = false;
            }
        }

        TextFound = found;
        foundRange = found ? _currentRange : null;
        return found;
    }

    // Replace the search text with new text
    public void Replace(string replaceText)
    {
        if (TextFound) // Check if the text was found
        {
            _currentRange.Text = replaceText;
        }
        else
        {
            throw new ArgumentException($"Text '{_searchText}' not found in the current search range.");
        }
    }
}
