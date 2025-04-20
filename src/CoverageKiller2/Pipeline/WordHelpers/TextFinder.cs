using CoverageKiller2.DOM;
using CoverageKiller2.Logging;
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
        private readonly CKRange _selectedRange = default; // Range to search for text
        private CKRange _lastFoundRange = default;
        private readonly string _searchText; // Text to search for

        public string SearchText => Tracer.Trace(_searchText);

        /// <summary>
        /// Gets a value indicating whether the text was found in the current range.
        /// </summary>
        public bool TextFound
        {
            get => Tracer.Trace(_lastFoundRange != null);
        }

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
        public TextFinder(CKDocument ckDoc, string searchText, CKRange searchWithinRange = null)
        {
            Tracer.Enabled = true;
            if (ckDoc == null) throw new ArgumentNullException(nameof(ckDoc));
            if (string.IsNullOrWhiteSpace(searchText)) throw new ArgumentNullException(nameof(searchText));



            _ckDoc = ckDoc;
            _selectedRange = searchWithinRange ?? _ckDoc.Content; // Use the entire document if no range is provided
            _searchText = searchText;

            Tracer.Log("Initialized", new DataPoints()
                .Add(nameof(CKDocument), ckDoc.FullPath)
                .Add(nameof(SearchText), SearchText)
                .Add($"{nameof(_selectedRange)}.Start", _selectedRange.Start)
                .Add($"{nameof(_selectedRange)}.End", _selectedRange.End)
                .Add($"{nameof(_lastFoundRange)}[Is Null]", _lastFoundRange == null));
        }



        public Tracer Tracer = new Tracer(typeof(TextFinder));

        ///// <summary>
        ///// Attempts to find the next occurrence of the search text within the specified range.
        ///// If it's the first call, it searches for the first match.
        ///// </summary>
        ///// <param name="foundRange">Outputs the found range if the text is found; otherwise, null.</param>
        ///// <param name="wrap">Indicates whether to wrap around to the beginning of the range if the end is reached.</param>
        ///// <returns>True if the text is found; otherwise, false.</returns>
        //public bool TryFind(out Word.Range foundRange, bool wrap = false)
        //{
        //    Tracer.Log("Finding...", new DataPoints()
        //        .Add(nameof(wrap) + "(inop)", wrap)
        //        .Add(nameof(SearchText), SearchText)
        //        .Add($"{nameof(_selectedRange)}.Start", _selectedRange?.Start)
        //        .Add($"{nameof(_selectedRange)}.End", _selectedRange?.End)
        //        .Add($"{nameof(_lastFoundRange)}[Is Null]", _lastFoundRange == null));


        //    var searchRangeStart = _lastFoundRange is null
        //        ? _selectedRange.Start
        //        : _lastFoundRange.End + 1; // not safe for end of document.

        //    var activeSearchRange = _ckDoc.Range(searchRangeStart, _selectedRange.End);

        //    _lastFoundRange = null;

        //    // Try to find the next occurrence
        //    var found = activeSearchRange.TryFindNext(SearchText, matchWildcards: true);

        //    if (found != null)
        //    {
        //        _lastFoundRange = activeSearchRange.CKCopy();
        //    }

        //    foundRange = _lastFoundRange is null ? null : _lastFoundRange.CKCopy();

        //    activeSearchRange = null;

        //    Tracer.Log($"Text found: {(found ? foundRange.Text : "[NO TEXT FOUND]")}",
        //        new DataPoints()
        //        .Add($"{nameof(_selectedRange)}.Start", _selectedRange?.Start)
        //        .Add($"{nameof(_selectedRange)}.End", _selectedRange?.End)
        //        .Add($"{nameof(_lastFoundRange)}.Start", _lastFoundRange?.Start)
        //        .Add($"{nameof(_lastFoundRange)}.End", _lastFoundRange?.End)
        //    );

        //    return found;
        //}


        public void Replace(string replaceText)
        {
            Tracer.Log("Replacing Text...", new DataPoints()
                .Add(nameof(replaceText), replaceText)
                .Add($"{nameof(_lastFoundRange)}.Start", _lastFoundRange?.Start)
                .Add($"{nameof(_lastFoundRange)}.End", _lastFoundRange?.End)
            );

            if (TextFound)
            {
                _lastFoundRange.Text = replaceText;

                _lastFoundRange = _ckDoc.Range(
                    _lastFoundRange.Start,
                    _lastFoundRange.Start + replaceText.Length);
            }
        }
    }
}
public static class WordRangeExtensions
{

    public static Word.Range CKCopy(this Word.Range source)
    {
        // Create a new range based on the source range's start and end
        Word.Range copiedRange = source.Document.Range(source.Start, source.End);
        return copiedRange;
    }
}