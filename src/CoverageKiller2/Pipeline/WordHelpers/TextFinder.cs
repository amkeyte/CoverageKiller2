using System;
using Word = Microsoft.Office.Interop.Word;

namespace CoverageKiller2.Pipeline.WordHelpers
{
    /// <summary>
    /// A utility class for finding and replacing text in a Word document.
    /// </summary>
    public class TextFinder
    {
        private readonly CKDocument _ckDoc; // Wrapper for Word.Document
        private Word.Range _currentRange; // Current range to search within
        private readonly Word.Range _searchWithinRange; // Range to search for text
        private readonly string _searchText; // Text to search for

        /// <summary>
        /// Gets a value indicating whether the text was found in the current range.
        /// </summary>
        public bool TextFound { get; private set; } = false;

        /// <summary>
        /// Initializes a new instance of the <see cref="TextFinder"/> class with a specified CKDocument and search text.
        /// </summary>
        /// <param name="ckDoc">The CKDocument object that contains the Word document.</param>
        /// <param name="searchText">The text to search for in the document.</param>
        public TextFinder(CKDocument ckDoc, string searchText)
            : this(ckDoc, searchText, null)
        {
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="TextFinder"/> class with a specified CKDocument, search text, and a range to search within.
        /// </summary>
        /// <param name="ckDoc">The CKDocument object that contains the Word document.</param>
        /// <param name="searchText">The text to search for in the document.</param>
        /// <param name="searchWithinRange">The range to search within. If null, the entire document will be searched.</param>
        public TextFinder(CKDocument ckDoc, string searchText, Word.Range searchWithinRange)
        {
            _ckDoc = ckDoc;
            _searchWithinRange = searchWithinRange ?? _ckDoc.Content; // Use the entire document if no range is provided
            _currentRange = _searchWithinRange;
            _searchText = searchText;
        }

        /// <summary>
        /// Attempts to find the first occurrence of the search text within the specified range.
        /// </summary>
        /// <param name="foundRange">Outputs the found range if the text is found; otherwise, null.</param>
        /// <returns>True if the text is found; otherwise, false.</returns>
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

        /// <summary>
        /// Attempts to find the next occurrence of the search text within the specified range.
        /// </summary>
        /// <param name="foundRange">Outputs the found range if the text is found; otherwise, null.</param>
        /// <param name="wrap">Indicates whether to wrap around to the beginning of the range if the end is reached.</param>
        /// <returns>True if the next occurrence is found; otherwise, false.</returns>
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

        /// <summary>
        /// Replaces the found text with the specified replacement text.
        /// </summary>
        /// <param name="replaceText">The text to replace the found text with.</param>
        /// <exception cref="ArgumentException">Thrown when the search text was not found.</exception>
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

            TextFound = false;
        }
    }
}
