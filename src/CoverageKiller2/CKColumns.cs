using Serilog;
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
            var xxx = new CKCells(((Word.Table)columns.Parent).Range.Cells);

            Log.Debug("TRACE => {class}.{func}() = {pVal1}",
                nameof(CKColumns),
                "ctor",
                $"{nameof(columns)}[(Table)Columns.Parent.{nameof(xxx.ContainsMerged)} = {xxx.ContainsMerged}]");

            _columns = columns ?? throw Crash.LogThrow(
                new ArgumentNullException(nameof(columns)));

            //cant use CKTable because no index.

            if (xxx.ContainsMerged)
            {
                throw Crash.LogThrow(
                    new InvalidOperationException("Cannot access individual columns in this collection because the table has mixed cell widths."));
            }

            try
            {
                _ = columns.Cast<Word.Column>().Any();
            }
            catch (Exception ex)
            {
                throw Crash.LogThrow(ex);
            }
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
                CKColumn col = default;
                try
                {
                    col = new CKColumn(_columns[i]);
                }
                catch (Exception ex)
                {
                    throw Crash.LogThrow(
                        new InvalidOperationException($"How did you get here? This object should not init?", ex));
                }
                yield return col;
            }
        }

        // Non-generic IEnumerable implementation
        IEnumerator IEnumerable.GetEnumerator()
        {
            return GetEnumerator();
        }

    }
}
