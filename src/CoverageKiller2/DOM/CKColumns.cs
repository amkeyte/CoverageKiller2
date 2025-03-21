using System;
using System.Collections;
using System.Collections.Generic;
using Word = Microsoft.Office.Interop.Word;

namespace CoverageKiller2.DOM
{
    public class CKColumns : IEnumerable<CKColumn>
    {
        public static CKColumns Create(CKRange parent)
        {
            return new CKColumns(parent);
        }
        internal Word.Columns COMObject => Parent.COMRange.Columns;

        public CKColumns(CKRange parent)
        {
            Parent = parent ?? throw new ArgumentNullException(nameof(parent));
        }


        // Property to get the total number of columns
        public int Count => COMObject.Count;

        public CKRange Parent { get; private set; }

        // Access a CKColumn by its current index (1-based index in Word)
        public CKColumn this[int index]
        {
            get
            {
                if (index < 1 || index > COMObject.Count)
                    throw new ArgumentOutOfRangeException(nameof(index), "Index must be between 1 and Count.");

                return CKColumn.Create(this, index);
            }
        }

        // IEnumerable implementation to allow foreach enumeration
        public IEnumerator<CKColumn> GetEnumerator()
        {
            for (int index = 1; index <= Count; index++)
            {
                yield return this[index];
            }
        }

        // Non-generic IEnumerable implementation
        IEnumerator IEnumerable.GetEnumerator()
        {
            return GetEnumerator();
        }

    }
}
