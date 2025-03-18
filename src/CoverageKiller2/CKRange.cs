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
        /// Gets the starting position of the range.
        /// </summary>
        public int Start => COMObject.Start;

        /// <summary>
        /// Gets the ending position of the range.
        /// </summary>
        public int End => COMObject.End;

        /// <summary>
        /// Initializes a new instance of the <see cref="CKRange"/> class.
        /// </summary>
        /// <param name="range">The Word.Range object to wrap.</param>
        /// <exception cref="ArgumentNullException">Thrown when the <paramref name="range"/> is null.</exception>
        private CKRange(Word.Range range)
        {
            COMObject = range ?? throw new ArgumentNullException(nameof(range));
        }

        /// <summary>
        /// Creates a new instance of <see cref="CKRange"/> wrapping the specified Word.Range.
        /// </summary>
        /// <param name="range">The Word.Range object to wrap.</param>
        /// <returns>A new instance of <see cref="CKRange"/>.</returns>
        internal static CKRange Create(Word.Range range)
        {
            return new CKRange(range);
        }
    }
}
