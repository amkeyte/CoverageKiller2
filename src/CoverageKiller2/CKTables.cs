using System;
using System.Collections;
using System.Collections.Generic;
using Word = Microsoft.Office.Interop.Word;

namespace CoverageKiller2
{
    /// <summary>
    /// Represents a collection of <see cref="CKTable"/> objects associated with a <see cref="CKRange"/>.
    /// </summary>
    public class CKTables : IEnumerable<CKTable>
    {
        /// <summary>
        /// Creates a new instance of <see cref="CKTables"/> for the specified <see cref="CKRange"/>.
        /// </summary>
        /// <param name="parent">The parent <see cref="CKRange"/> that contains the tables.</param>
        /// <returns>A new instance of <see cref="CKTables"/>.</returns>
        /// <exception cref="ArgumentNullException">Thrown if <paramref name="parent"/> is null.</exception>
        internal static CKTables Create(CKRange parent)
        {
            parent = parent ?? throw new ArgumentNullException(nameof(parent));
            return new CKTables(parent);
        }

        /// <summary>
        /// Gets the parent <see cref="CKRange"/> associated with this instance.
        /// </summary>
        internal CKRange Parent { get; private set; }

        /// <summary>
        /// Gets the underlying Word.Tables COM object from the parent range.
        /// Note that there is only one Tables property, so calling back to it
        /// instead of storing a reference every time is acceptable.
        /// </summary>
        internal Word.Tables COMObject => Parent.COMObject.Tables;

        /// <summary>
        /// Returns a string that represents the current <see cref="CKTables"/> instance.
        /// </summary>
        /// <returns>A string containing the count of tables.</returns>
        public override string ToString()
        {
            // Since CKRange doesn't provide a file path, we simply return the count.
            return $"CKTables [Count: {Count}]";
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="CKTables"/> class.
        /// </summary>
        /// <param name="parent">The parent <see cref="CKRange"/> to associate with this instance.</param>
        private CKTables(CKRange parent)
        {
            Parent = parent;
        }

        /// <summary>
        /// Gets the number of tables in the associated range.
        /// </summary>
        public int Count => COMObject.Count;

        /// <summary>
        /// Gets the <see cref="CKTable"/> at the specified one-based index.
        /// </summary>
        /// <param name="index">The one-based index of the table to retrieve.</param>
        /// <returns>The <see cref="CKTable"/> at the specified index.</returns>
        /// <exception cref="ArgumentOutOfRangeException">
        /// Thrown when the index is less than 1 or greater than the number of tables.
        /// </exception>
        public CKTable this[int index]
        {
            get
            {
                if (index < 1 || index > Count)
                {
                    throw new ArgumentOutOfRangeException(nameof(index), "Index must be between 1 and the number of tables.");
                }
                return CKTable.Create(this, index);
            }
        }

        /// <summary>
        /// Determines the one-based index of the specified <see cref="CKTable"/> in the collection.
        /// </summary>
        /// <param name="targetTable">The table to locate in the collection.</param>
        /// <returns>
        /// The one-based index of the table if found; otherwise, -1.
        /// </returns>
        public int IndexOf(CKTable targetTable)
        {
            for (int i = 1; i <= Count; i++)
            {
                var table = COMObject[i];

                // Compare by checking that both tables have the same start and end range
                if (table.Range.Start == targetTable.COMObject.Range.Start &&
                    table.Range.End == targetTable.COMObject.Range.End)
                {
                    return i;
                }
            }
            return -1;
        }

        /// <summary>
        /// Returns an enumerator that iterates through the <see cref="CKTable"/> objects in the collection.
        /// </summary>
        /// <returns>An enumerator for the collection of <see cref="CKTable"/> objects.</returns>
        public IEnumerator<CKTable> GetEnumerator()
        {
            for (int i = 1; i <= Count; i++)
            {
                yield return this[i];
            }
        }

        /// <summary>
        /// Returns an enumerator that iterates through the collection.
        /// </summary>
        /// <returns>An enumerator for the collection.</returns>
        IEnumerator IEnumerable.GetEnumerator() => GetEnumerator();

        /// <summary>
        /// This method is not implemented.
        /// </summary>
        /// <returns>Throws a <see cref="NotImplementedException"/> when called.</returns>
        internal static object ToList()
        {
            throw new NotImplementedException();
        }
    }
}
