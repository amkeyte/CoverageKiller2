using CoverageKiller2.Tests;  // Contains LiveWordDocument helper.
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.Collections.Generic;
using System.Linq;
using Word = Microsoft.Office.Interop.Word;

namespace CoverageKiller2.DOM
{
    [TestClass]
    public class CKRowsTests
    {
        // Constant for the test table number (one-based). Adjust if needed.
        private const int TestTableNumber = 1;

        /// <summary>
        /// Tests that the CKRows collection count matches the number of rows in the live Word table.
        /// </summary>
        [TestMethod]
        public void CKRows_Count_MatchesWordTableRows()
        {
            LiveWordDocument.WithTestDocument(LiveWordDocument.Default, doc =>
            {
                // Verify the document contains at least TestTableNumber tables.
                Assert.IsTrue(doc.Tables.Count >= TestTableNumber,
                    $"Document must contain at least {TestTableNumber} table(s).");

                Word.Table wordTable = doc.Tables[TestTableNumber];
                CKTable ckTable = new CKTable(wordTable);

                // Build a list of CKRow objects from each row of the table.
                var rows = new List<CKRow>();
                int tableRowCount = wordTable.Rows.Count;
                int tableColumnCount = wordTable.Columns.Count;
                for (int i = 1; i <= tableRowCount; i++)
                {
                    // Create a rectangular cell reference for the entire row.
                    var cellRef = CKCellRefRect.ForRow(i, tableColumnCount);
                    // Create a CKRow using the table and its cell reference.
                    CKRow row = new CKRow(ckTable, cellRef);
                    rows.Add(row);
                }

                // Construct the CKRows collection using the list of CKRow objects.
                CKRows ckRows = new CKRows(rows, ckTable);

                // Verify that the CKRows count equals the number of rows in the Word table.
                Assert.AreEqual(tableRowCount, ckRows.Count,
                    "The CKRows count should match the Word table's row count.");
            });
        }

        /// <summary>
        /// Tests that the indexer returns the correct CKRow and that each row's Index property is correct.
        /// </summary>
        [TestMethod]
        public void CKRows_Indexer_ReturnsCorrectRowIndex()
        {
            LiveWordDocument.WithTestDocument(LiveWordDocument.Default, doc =>
            {
                Assert.IsTrue(doc.Tables.Count >= TestTableNumber,
                    $"Document must contain at least {TestTableNumber} table(s).");

                Word.Table wordTable = doc.Tables[TestTableNumber];
                CKTable ckTable = new CKTable(wordTable);

                int tableRowCount = wordTable.Rows.Count;
                int tableColumnCount = wordTable.Columns.Count;
                var rows = new List<CKRow>();
                for (int i = 1; i <= tableRowCount; i++)
                {
                    var cellRef = CKCellRefRect.ForRow(i, tableColumnCount);
                    CKRow row = new CKRow(ckTable, cellRef);
                    rows.Add(row);
                }

                CKRows ckRows = new CKRows(rows, ckTable);

                // Test one-based indexing: the first row should have index 1, etc.
                for (int i = 1; i <= ckRows.Count; i++)
                {
                    Assert.AreEqual(i, ckRows[i].Index, $"Row at one-based index {i} should have Index {i}.");
                }
            });
        }

        /// <summary>
        /// Tests that enumerating the CKRows collection yields the correct number of rows.
        /// </summary>
        [TestMethod]
        public void CKRows_Enumeration_YieldsAllRows()
        {
            LiveWordDocument.WithTestDocument(LiveWordDocument.Default, doc =>
            {
                Assert.IsTrue(doc.Tables.Count >= TestTableNumber,
                    $"Document must contain at least {TestTableNumber} table(s).");

                Word.Table wordTable = doc.Tables[TestTableNumber];
                CKTable ckTable = new CKTable(wordTable);

                int tableRowCount = wordTable.Rows.Count;
                int tableColumnCount = wordTable.Columns.Count;
                var rows = new List<CKRow>();
                for (int i = 1; i <= tableRowCount; i++)
                {
                    var cellRef = CKCellRefRect.ForRow(i, tableColumnCount);
                    CKRow row = new CKRow(ckTable, cellRef);
                    rows.Add(row);
                }

                CKRows ckRows = new CKRows(rows, ckTable);

                int enumeratedCount = ckRows.Count();
                Assert.AreEqual(ckRows.Count, enumeratedCount,
                    "Enumeration of CKRows should yield the same number of rows as the Count property.");
            });
        }
    }
}
