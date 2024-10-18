using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using Word = Microsoft.Office.Interop.Word;

namespace CoverageKiller2
{
    public class CKColumns : IEnumerable<CKColumn>
    {
        private Word.Columns _columns;
        // Constructor to initialize CKColumns with Word.Columns
        public CKColumns(Word.Columns columns)
        {
            _columns = columns ?? throw new ArgumentNullException(nameof(columns));
        }

        public bool ContainsMerged => _columns.Cast<Word.Column>()
            .Any(col => new CKColumn(col).ContainsMerged);

        // Property to get the total number of columns
        public int Count => _columns.Count;

        // Access a CKColumn by its index (1-based index in Word)
        public CKColumn this[int index]
        {
            get
            {
                if (index < 1 || index > _columns.Count)
                    throw new ArgumentOutOfRangeException(nameof(index), "Index must be between 1 and Count.");

                return new CKColumn(_columns[index]);
            }
        }

        // IEnumerable implementation to allow foreach enumeration
        public IEnumerator<CKColumn> GetEnumerator()
        {
            for (int i = 1; i <= _columns.Count; i++)
            {
                yield return new CKColumn(_columns[i]);
            }
        }

        // Non-generic IEnumerable implementation
        IEnumerator IEnumerable.GetEnumerator()
        {
            return GetEnumerator();
        }

    }
}
