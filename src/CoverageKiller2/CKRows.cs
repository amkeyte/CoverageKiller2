using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using Word = Microsoft.Office.Interop.Word;

namespace CoverageKiller2
{
    public class CKRows : IEnumerable<CKRow>
    {
        private Word.Rows _rows;

        // Constructor to initialize CKRows with Word.Rows
        public CKRows(Word.Rows rows)
        {
            _rows = rows ?? throw new ArgumentNullException(nameof(rows));
        }

        public bool ContainsMerged => _rows.Cast<Word.Row>()
            .Any(row => new CKRow(row).ContainsMerged);
        // Property to get the total number of rows
        public int Count => _rows.Count;

        // Access a CKRow by its index (1-based index in Word)
        public CKRow this[int index]
        {
            get
            {
                if (index < 1 || index > _rows.Count)
                    throw new ArgumentOutOfRangeException(nameof(index), "Index must be between 1 and Count.");

                return new CKRow(_rows[index]);
            }
        }

        // IEnumerable implementation to allow foreach enumeration
        public IEnumerator<CKRow> GetEnumerator()
        {
            for (int i = 1; i <= _rows.Count; i++)
            {
                yield return new CKRow(_rows[i]);
            }
        }

        // Non-generic IEnumerable implementation
        IEnumerator IEnumerable.GetEnumerator()
        {
            return GetEnumerator();
        }

        // Additional helper method to get the first row
        public CKRow First()
        {
            if (_rows.Count == 0)
                throw new InvalidOperationException("No rows in the collection.");

            return new CKRow(_rows[1]);
        }

        // Additional helper method to get the last row
        public CKRow Last()
        {
            if (_rows.Count == 0)
                throw new InvalidOperationException("No rows in the collection.");

            return new CKRow(_rows[_rows.Count]);
        }
    }
}
