using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using Word = Microsoft.Office.Interop.Word;

///************
///Yes, it behaves similarly. When you access the Tables property on a Word.Range, Word returns a collection of all tables that intersect with that range. This means:

//If the range is entirely within a table, you'll get that table.
//If the range spans across multiple tables, you'll get all tables that the range touches.
//If the range doesn’t include any tables, the collection will simply be empty.
//No error is thrown if the range is smaller than a table or doesn’t fully encompass one.
///************


namespace CoverageKiller2.DOM
{
    /// <summary>
    /// Represents a collection of <see cref="CKTable"/> objects associated with a <see cref="CKRange"/>.
    /// </summary>
    public class CKTables : ACKRangeCollection, IEnumerable<CKTable>
    {
        /// <summary>
        /// Gets the underlying Word.Tables COM object from the parent range.
        /// Note that there is only one Tables property, so calling back to it
        /// instead of storing a reference every time is acceptable.
        /// </summary>
        internal Word.Tables COMTables => Parent.COMRange.Tables;

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
        internal CKTables(CKRange parent) : base(parent) { }

        /// <summary>
        /// Gets the number of tables in the associated range.
        /// </summary>
        public override int Count => COMTables.Count;

        /// <summary>
        /// optimize this if there's a big delay here.
        /// </summary>
        public override bool IsDirty => _isDirty || this.Any(x => x.IsDirty);

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
                return new CKTable(COMTables[index]);
            }
        }

        /// <summary>
        /// Determines the one-based index of the specified <see cref="CKTable"/> in the collection.
        /// </summary>
        /// <param name="targetTable">The table to locate in the collection.</param>
        /// <returns>
        /// The one-based index of the table if found; otherwise, -1.
        /// </returns>
        //public int IndexOf(CKTable targetTable)
        //{
        //    for (int i = 1; i <= Count; i++)
        //    {
        //        var table = COMTables[i];

        //        // Compare by checking that both tables have the same start and end range
        //        if (table.Range.Start == targetTable.COMObject.Range.Start &&
        //            table.Range.End == targetTable.COMObject.Range.End)
        //        {
        //            return i;
        //        }
        //    }
        //    return -1;
        //}

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
    }
}
