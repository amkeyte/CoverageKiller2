using System;
using System.Collections;
using System.Collections.Generic;
using Word = Microsoft.Office.Interop.Word;

namespace CoverageKiller2
{
    public abstract class CKCells : IEnumerable<CKCell>
    {
        public static CKCells Create(CKColumn parent)
        {
            return new CKColumnCells(parent);
        }

        public static CKCells Create(CKRow parent)
        {
            return new CKRowCells(parent);
        }
        internal static CKCells Create(CKTable parent)
        {
            return new CKTableCells(parent);
        }



        public abstract Word.Cells COMObject { get; }
        public int Count => COMObject.Count;
        public CKCell this[int index]
        {
            get
            {
                if (index < 1 || index > Count)
                    throw new ArgumentOutOfRangeException(nameof(index), "Index must be between 1 and Count.");

                return CKCell.Create(this, index);// new CKCell(COMObject[index]);
            }
        }

        public IEnumerator<CKCell> GetEnumerator()
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
    }
    internal class CKColumnCells : CKCells
    {
        public CKColumnCells(CKColumn parent)
        {
            Parent = parent;
        }

        public override Word.Cells COMObject => Parent.COMObject.Cells;
        public CKColumn Parent { get; set; }

    }
    internal class CKRowCells : CKCells
    {
        public CKRowCells(CKRow parent)
        {
            Parent = parent;
        }

        public override Word.Cells COMObject => Parent.COMObject.Cells;
        public CKRow Parent { get; private set; }
        public int Count => COMObject.Count;
    }
    internal class CKTableCells : CKCells
    {
        public CKTableCells(CKTable parent)
        {
            Parent = parent;
        }

        public override Word.Cells COMObject => Parent.COMObject.Range.Cells;
        public CKTable Parent { get; set; }
    }



}
