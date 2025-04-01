using CoverageKiller2.DOM;
using CoverageKiller2.DOM.Tables;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using Word = Microsoft.Office.Interop.Word;

namespace CoverageKiller2.Tests.Tables
{
    public static class TableTestHelpers
    {
        /// <summary>
        /// Returns all CKTables in the document that contain merged cells.
        /// </summary>
        public static IReadOnlyList<CKTable> GetAllMergedTables(CKDocument doc)
        {
            if (doc == null) throw new ArgumentNullException(nameof(doc));

            var mergedTables = new List<CKTable>();

            foreach (var ckTable in doc.Tables)
            {
                Word.Table wordTable = ckTable.COMTable;

                foreach (Word.Cell cell in wordTable.Range.Cells)
                {
                    if (IsMerged(cell))
                    {
                        mergedTables.Add(ckTable);
                        Debug.WriteLine($"Table: {mergedTables.Count} \"{ckTable.Cell(1).Text}\"");
                        break; // This table has at least one merged cell
                    }
                }
            }

            return mergedTables;
        }

        /// <summary>
        /// Returns the Nth CKTable in the document that contains merged cells.
        /// </summary>
        public static CKTable FindNthMergedTable(CKDocument doc, int index)
        {
            var mergedTables = GetAllMergedTables(doc);

            if (index < 1 || index > mergedTables.Count)
                throw new ArgumentOutOfRangeException(nameof(index), $"Document contains only {mergedTables.Count} merged tables.");

            return mergedTables[index - 1];
        }

        /// <summary>
        /// Determines whether a given Word cell is part of a merged region.
        /// </summary>
        private static bool IsMerged(Word.Cell cell)
        {
            try
            {
                // Merged cells usually span multiple logical cells
                return cell.Range.Cells.Count > 1;
            }
            catch
            {
                return false;
            }
        }
    }
}
