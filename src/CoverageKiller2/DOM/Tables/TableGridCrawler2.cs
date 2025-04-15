using System;
using System.Linq;
using Word = Microsoft.Office.Interop.Word;

namespace CoverageKiller2.DOM.Tables
{
    internal class TableGridCrawler2
    {

        public static Base1JaggedList<Word.Cell> GetVisualRows(Word.Table table)
        {
            return new Base1JaggedList<Word.Cell>(table.Range.Cells
                .ToList()
                .GroupBy(c => c.RowIndex)
                .OrderBy(g => g.Key)
                .Select(g => g.OrderBy(c => c.ColumnIndex).ToList())
                .ToList());
        }

        public static void NormalizeJaggedList(Base1JaggedList<Word.Cell> CellsList)
        {
            int maxRowCount = 0;

            //prefix/insert
            foreach (var row in CellsList)
            {

                var rowCopy = new Base1List<Word.Cell>(row);
                row.Clear();
                Word.Cell lastCell = null;

                foreach (var wordCell in rowCopy)
                {
                    if (wordCell is null) continue;

                    int padColAmount = 0;
                    if (lastCell != null)
                    {
                        padColAmount = wordCell.ColumnIndex - lastCell.ColumnIndex - 1;
                    }
                    else
                    {
                        padColAmount = wordCell.ColumnIndex - 1;
                    }

                    while (padColAmount-- > 0)
                    {
                        row.Add(null);
                    }

                    row.Add(wordCell);
                    maxRowCount = Math.Max(maxRowCount, row.Count);
                    lastCell = wordCell;
                }



            }

            //int rowsToFix = 1;
            //while (rowsToFix > 0)//safety
            //{
            var x = CellsList.Where(r => r.Count < maxRowCount);
            //rowsToFix = x.Count();


        }
    }
}
