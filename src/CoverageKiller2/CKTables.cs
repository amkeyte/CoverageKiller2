using System;
using System.Collections;
using System.Collections.Generic;
using Word = Microsoft.Office.Interop.Word;

namespace CoverageKiller2
{
    public class CKTables : IEnumerable<CKTable>
    {
        private readonly Word.Tables _tables;

        public CKTables(Word.Tables tables)
        {
            _tables = tables ?? throw new ArgumentNullException(nameof(tables));
        }

        public int Count => _tables.Count;

        public CKTable this[int index]
        {
            get
            {
                if (index < 1 || index > Count)
                {
                    throw new ArgumentOutOfRangeException(nameof(index), "Index must be between 1 and the number of tables.");
                }
                return new CKTable(_tables[index]); // Assuming CKTable wraps a Word.Table instance
            }
        }

        public IEnumerator<CKTable> GetEnumerator()
        {
            for (int i = 1; i <= Count; i++)
            {
                yield return this[i];
            }
        }

        IEnumerator IEnumerable.GetEnumerator() => GetEnumerator();

    }
}
