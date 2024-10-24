using System;
using System.Collections;
using System.Collections.Generic;
using Word = Microsoft.Office.Interop.Word;

namespace CoverageKiller2
{
    public class CKTables : IEnumerable<CKTable>
    {
        internal static CKTables Create(CKDocument parent)
        {
            parent = parent ?? throw new ArgumentNullException(nameof(parent));
            return new CKTables(parent);
        }
        internal CKDocument Parent { get; private set; }

        //there is only one Tables property, so calling back to it instead of
        //storing a reference every time is fine.
        internal Word.Tables COMObject => Parent.COMObject.Tables;

        private CKTables(CKDocument parent)
        {
            Parent = parent;

        }

        public int Count => COMObject.Count;

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
