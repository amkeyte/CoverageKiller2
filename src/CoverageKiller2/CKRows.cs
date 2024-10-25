using System;
using System.Collections;
using System.Collections.Generic;
using Word = Microsoft.Office.Interop.Word;

namespace CoverageKiller2
{
    public class CKRows : IEnumerable<CKRow>
    {
        public static CKRows Create(CKTable parent)
        {
            return new CKRows(parent);
        }
        internal Word.Rows COMObject => Parent.COMObject.Rows;

        public CKTable Parent { get; private set; }


        public override string ToString()
        {
            //string docName = System.IO.Path.GetFileName(Parent.FullPath);

            return $"Table[{Parent._lastIndex}].Rows[Count: {Count}]";

        }





        // Constructor to initialize CKRows with Word.Rows
        public CKRows(CKTable parent)
        {
            Parent = parent;
        }

        //public bool ContainsMerged => _rows.Cast<Word.Row>()
        //    .Any(row => new CKRow(row).ContainsMerged);
        // Property to get the total number of rows
        public int Count => COMObject.Count;

        // Access a CKRow by its index (1-based index in Word)
        public CKRow this[int index]
        {
            get
            {
                if (index < 1 || index > Count)
                    throw new ArgumentOutOfRangeException(nameof(index), "Index must be between 1 and Count.");

                return CKRow.Create(this, index);
            }
        }

        // IEnumerable implementation to allow foreach enumeration
        public IEnumerator<CKRow> GetEnumerator()
        {
            for (int i = 1; i <= Count; i++)
            {
                yield return this[i];
            }
        }

        // Non-generic IEnumerable implementation
        IEnumerator IEnumerable.GetEnumerator()
        {
            return GetEnumerator();
        }

        //// Additional helper method to get the first row
        //public CKRow First()
        //{
        //    if (Count == 0)
        //        throw new InvalidOperationException("No rows in the collection.");

        //    return new CKRow.Create(this,1);
        //}

        //// Additional helper method to get the last row
        //public CKRow Last()
        //{
        //    if (_rows.Count == 0)
        //        throw new InvalidOperationException("No rows in the collection.");

        //    return new CKRow(_rows[_rows.Count]);
        //}

        internal CKRow Add(CKRow beforeRow)
        {
            //Add inserts
            return CKRow.Create(this, COMObject.Add(beforeRow.COMObject).Index);
        }
    }
}
