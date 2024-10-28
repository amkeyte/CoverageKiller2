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
        public override string ToString()
        {
            string docName = System.IO.Path.GetFileName(Parent.FullPath);

            return $"[{docName}].Tables[Count: {Count}]";

        }
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

            // Return -1 if the target table is not found
            return -1;
        }
        public IEnumerator<CKTable> GetEnumerator()
        {
            for (int i = 1; i <= Count; i++)
            {
                yield return this[i];
            }
        }

        IEnumerator IEnumerable.GetEnumerator() => GetEnumerator();

        internal static object ToList()
        {
            throw new NotImplementedException();
        }
    }
}
