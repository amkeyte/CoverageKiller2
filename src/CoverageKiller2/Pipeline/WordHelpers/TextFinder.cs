using Serilog;
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

        public string SearchText => _searchText;

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

            Log.Debug(LH.TraceCaller(LH.PP.Result, "Values Initialized",
                nameof(TextFinder), "ctor",
                nameof(CKDocument), ckDoc.FullPath,
                nameof(searchText), searchText,
                nameof(_searchWithinRange.Start), _searchWithinRange.Start,
                nameof(_searchWithinRange.End), _searchWithinRange.End));


        }

        /// <summary>
        /// Attempts to find the next occurrence of the search text within the specified range.
        /// If it's the first call, it searches for the first match.
        /// </summary>
        /// <param name="foundRange">Outputs the found range if the text is found; otherwise, null.</param>
        /// <param name="wrap">Indicates whether to wrap around to the beginning of the range if the end is reached.</param>
        /// <returns>True if the text is found; otherwise, false.</returns>
        public bool TryFind(out Word.Range foundRange, bool wrap = false)
        {
            Log.Debug(LH.TraceCaller(LH.PP.Enter, null,
                nameof(TextFinder), nameof(TryFind),
                nameof(wrap), wrap));

            Log.Debug(LH.TraceCaller(LH.PP.Result, "Test Point; first find.",
                nameof(_searchText), _searchText,
                nameof(_searchWithinRange.Start), _searchWithinRange.Start.ToString(),
                nameof(_searchWithinRange.End), _searchWithinRange.End.ToString()));


            TextFound = false;
            int originalStart = _currentRange.Start; // Save the original starting point

            // Move range start to after the last found occurrence
            //_currentRange.Start = _currentRange.End;

            // Try to find the next occurrence
            bool found = _currentRange.Find.Execute(
                FindText: _searchText,
                MatchWildcards: true);

            //Log.Debug(LH.TraceCaller(
            //nameof(TextFinder), "TryFind- after first find",
            //nameof(_searchText), _searchText,
            //nameof(found), found.ToString(),
            //nameof(_searchWithinRange.Start), _searchWithinRange.Start.ToString(),
            //nameof(_searchWithinRange.End), _searchWithinRange.End.ToString()));

            // If not found and wrapping is enabled, start from the beginning of the range
            //if (!found && wrap)
            //{
            //    _currentRange = _searchWithinRange; // Reset to the search range
            //    found = _currentRange.Find.Execute(FindText: _searchText, MatchWildcards: true);

            //    // Ensure we don't match the same initial occurrence (avoid infinite loop)
            //    if (found && _currentRange.Start == originalStart)
            //    {
            //        found = false;
            //    }

            //    Log.Debug(LH.TraceCaller(
            //    nameof(TextFinder), "TryFind- after wrap",
            //    nameof(_searchText), _searchText,
            //    nameof(found), found.ToString(),
            //    nameof(_searchWithinRange.Start), _searchWithinRange.Start.ToString(),
            //    nameof(_searchWithinRange.End), _searchWithinRange.End.ToString()));
            //}

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

            TextFound = false; // Reset TextFound after replacement
        }
    }
}
