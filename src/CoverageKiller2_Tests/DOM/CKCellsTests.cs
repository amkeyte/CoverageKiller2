using CoverageKiller2.Tests;  // Contains LiveWordDocument helper.
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Linq;
using Word = Microsoft.Office.Interop.Word;

namespace CoverageKiller2.DOM
{
    [TestClass]
    public class CKCellsTests
    {
        /// <summary>
        /// Tests that the CKCells collection constructed from a CKRange returns a valid cell collection.
        /// </summary>
        [TestMethod]
        public void CKCells_Constructor_FromRange_ReturnsValidCollection()
        {
            LiveWordDocument.WithTestDocument(LiveWordDocument.Default, doc =>
            {
                // Ensure the document contains at least one table.
                Assert.IsTrue(doc.Tables.Count > 0, "Document must contain at least one table.");

                // Use table 1 for testing.
                Word.Table wordTable = doc.Tables[1];

                // Create a CKRange from the table's range.
                CKRange range = new CKRange(wordTable.Range);

                // Construct CKCells from the range.
                CKCells cells = new CKCellsLinear(range);

                // The cells collection should not be null and should contain at least one cell.
                Assert.IsNotNull(cells, "CKCells instance should not be null.");
                Assert.IsTrue(cells.Count > 0, "CKCells should contain at least one cell.");
            });
        }

        /// <summary>
        /// Tests that CKCells constructed from a CKTable and a linear cell reference returns the correct number of cells.
        /// </summary>
        [TestMethod]
        public void CKCells_Constructor_FromTableAndCellRef_ReturnsValidCollection()
        {
            LiveWordDocument.WithTestDocument(LiveWordDocument.Default, doc =>
            {
                // Use the first table.
                Assert.IsTrue(doc.Tables.Count > 0, "Document must contain at least one table.");
                Word.Table wordTable = doc.Tables[1];
                CKTable ckTable = new CKTable(wordTable);

                // Get the number of cells in the table's range (assumed one-based collection).
                int cellCount = wordTable.Range.Cells.Count;

                // Create a linear cell reference for all cells in the table.
                ICellRef cellRef = CKCellRefLinear.ForCells(1, cellCount);

                // Construct the CKCells collection.
                CKCells cells = new CKCellsLinear(ckTable, cellRef);

                // Verify that the cells collection count matches the table's cells count.
                Assert.AreEqual(cellCount, cells.Count, "The number of CKCells should match the table's cell count.");
            });
        }

        /// <summary>
        /// Tests the indexer for CKCells returns valid cells and throws for out-of-range indices.
        /// </summary>
        [TestMethod]
        public void CKCells_Indexer_ReturnsValidCell_And_ThrowsOnInvalidIndex()
        {
            LiveWordDocument.WithTestDocument(LiveWordDocument.Default, doc =>
            {
                Word.Table wordTable = doc.Tables[1];
                CKTable ckTable = new CKTable(wordTable);
                int cellCount = wordTable.Range.Cells.Count;
                ICellRef cellRef = CKCellRefLinear.ForCells(1, cellCount);
                CKCells cells = new CKCellsLinear(ckTable, cellRef);

                // Test valid index.
                CKCell firstCell = cells[1];
                Assert.IsNotNull(firstCell, "The indexer should return a valid CKCell for a valid index.");

                // Test index 0 (invalid).
                Assert.ThrowsException<ArgumentOutOfRangeException>(() =>
                {
                    var cell0 = cells[0];
                }, "Accessing index 0 should throw ArgumentOutOfRangeException.");

                // Test index greater than Count.
                Assert.ThrowsException<ArgumentOutOfRangeException>(() =>
                {
                    var cellOut = cells[cellCount + 1];
                }, "Accessing an index greater than Count should throw ArgumentOutOfRangeException.");
            });
        }

        /// <summary>
        /// Tests that enumerating CKCells yields the correct number of cells.
        /// </summary>
        [TestMethod]
        public void CKCells_Enumeration_YieldsAllCells()
        {
            LiveWordDocument.WithTestDocument(LiveWordDocument.Default, doc =>
            {
                Word.Table wordTable = doc.Tables[1];
                CKTable ckTable = new CKTable(wordTable);
                int cellCount = wordTable.Range.Cells.Count;
                ICellRef cellRef = CKCellRefLinear.ForCells(1, cellCount);
                CKCells cells = new CKCellsLinear(ckTable, cellRef);

                int enumeratedCount = cells.Count();
                Assert.AreEqual(cellCount, enumeratedCount, "Enumeration should yield the same number of cells as the Count property.");
            });
        }
    }
}
