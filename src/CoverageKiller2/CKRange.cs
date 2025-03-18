using System;
using Word = Microsoft.Office.Interop.Word;

namespace CoverageKiller2
{
    /// <summary>
    /// Represents a simple wrapper for the Word.Range object.
    /// </summary>
    public class CKRange
    {
        /// <summary>
        /// Gets the underlying Word.Range COM object.
        /// </summary>
        public Word.Range COMObject { get; private set; }

        /// <summary>
        /// Initializes a new instance of the <see cref="CKRange"/> class.
        /// </summary>
        /// <param name="range">The Word.Range object to wrap.</param>
        /// <exception cref="ArgumentNullException">Thrown when the <paramref name="range"/> is null.</exception>
        public CKRange(Word.Range range)
        {
            COMObject = range ?? throw new ArgumentNullException(nameof(range));
        }
    }
}
