using Serilog;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using Word = Microsoft.Office.Interop.Word;

namespace CoverageKiller2
{
    public class CKCells : IEnumerable<CKCell>
    {
        private Word.Cells _cells;


        // Constructor to initialize CKCells with Word.Cells
        public CKCells(Word.Cells cells)
        {
            _cells = cells ?? throw new ArgumentNullException(nameof(cells));
        }

        // Property to get the total number of cells
        public int Count => _cells.Count;

        public bool ContainsMerged
        {
            get
            {
                var isMerged = _cells.Cast<Word.Cell>().Any(c => c.IsMerged());

                Log.Debug("TRACE => {class}.{prop}.get()",
                nameof(CKCells),
                nameof(ContainsMerged),
               $"Return[{nameof(isMerged)} = {isMerged}]");

                return isMerged;
            }
        }



        // Access a CKCell by its index (1-based index in Word)
        public CKCell this[int index]
        {
            get
            {
                if (index < 1 || index > _cells.Count)
                    throw new ArgumentOutOfRangeException(nameof(index), "Index must be between 1 and Count.");

                return new CKCell(_cells[index]);
            }
        }

        // IEnumerable implementation to allow foreach enumeration
        public IEnumerator<CKCell> GetEnumerator()
        {
            for (int i = 1; i <= _cells.Count; i++)
            {
                yield return new CKCell(_cells[i]);
            }
        }

        // Non-generic IEnumerable implementation
        IEnumerator IEnumerable.GetEnumerator()
        {
            return GetEnumerator();
        }

    }
}
