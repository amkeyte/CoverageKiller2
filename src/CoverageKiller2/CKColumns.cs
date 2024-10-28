using CoverageKiller2.Logging;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using Word = Microsoft.Office.Interop.Word;

namespace CoverageKiller2
{
    public class CKColumns : IEnumerable<CKColumn>
    {
        public static CKColumns Create(CKTable parent)
        {
            return new CKColumns(parent);
        }
        internal Word.Columns COMObject => Parent.COMObject.Columns;

        public CKColumns(CKTable parent)
        {
            Parent = parent;
            //var xxx = new CKCells(((Word.Table)columns.Parent).Range.Cells);

            //Log.Debug("TRACE => {class}.{func}() = {pVal1}",
            //    nameof(CKColumns),
            //    "ctor",
            //    $"{nameof(columns)}[(Table)Columns.Parent.{nameof(xxx.ContainsMerged)} = {xxx.ContainsMerged}]");

            //_columns = columns ?? throw LH.LogThrow(
            //    new ArgumentNullException(nameof(columns)));

            //cant use CKTable because no index.

            //if (xxx.ContainsMerged)
            //{
            //    throw Crash.LogThrow(
            //        new InvalidOperationException("Cannot access individual columns in this collection because the table has mixed cell widths."));
            //}


            try
            {
                //hack for now until merged columns are supported.
                _ = COMObject.Cast<Word.Column>().Any();
            }
            catch (Exception ex)
            {
                throw LH.LogThrow(ex);
            }
        }


        // Property to get the total number of columns
        public int Count => COMObject.Count;

        public CKTable Parent { get; private set; }

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
