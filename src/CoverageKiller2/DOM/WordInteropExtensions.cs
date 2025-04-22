using System.Collections.Generic;
using System.Linq;
using Word = Microsoft.Office.Interop.Word;

namespace CoverageKiller2.DOM
{
    public static class WordInteropExtensions
    {
        public static IEnumerable<Word.Cell> AsEnumerable(this Word.Cells cells)
        {
            foreach (Word.Cell cell in cells)
            {
                yield return cell;
            }
        }
        public static List<Word.Cell> ToList(this Word.Cells cells)
        {
            return cells.AsEnumerable().ToList();

        }
        public static bool Contains(this Word.Range outer, Word.Range inner)
        {
            return inner.Start >= outer.Start && inner.End <= outer.End;
        }
        public static bool COMEquals(this Word.Range me, Word.Range other)
        {
            return me.Start >= other.Start && me.End <= other.End;
        }
    }
}

