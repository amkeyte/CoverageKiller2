using CoverageKiller2.Tests;  // Contains LiveWordDocument helper.
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections.Generic;
using System.Linq;
using Word = Microsoft.Office.Interop.Word;

namespace CoverageKiller2.DOM
{
    [TestClass]
    public class CKCellsTests
    {
        private class DummyRectRef : ICellRef<CKCellsRect>
        {
            public IEnumerable<int> WordCells => new[] { 1 };
            public int GridX1 { get; }
            public int GridY1 { get; }
            public int GridX2 { get; }
            public int GridY2 { get; }

            public DummyRectRef(int x1, int y1, int x2, int y2)
            {
                GridX1 = x1;
                GridY1 = y1;
                GridX2 = x2;
                GridY2 = y2;
            }
        }

        /// <summary>
        /// Tests that CKCells constructed from a CKTable and a rectangular cell reference returns the correct number of cells.
        /// </summary>
        [TestMethod]
        public void CKCells_Constructor_FromTableAndCellRef_ReturnsValidCollection()
        {
            LiveWordDocument.WithTestDocument(LiveWordDocument.Default, doc =>
            {
                Assert.IsTrue(doc.Tables.Count > 0, "Document must contain at least one table.");
                Word.Table wordTable = doc.Tables[1];
                CKTable ckTable = new CKTable(wordTable);

                var cellRef = new DummyRectRef(0, 0, 0, 0);
                CKCells cells = ckTable.Converters.GetCells(ckTable, ckTable, cellRef);

                Assert.IsNotNull(cells, "CKCells instance should not be null.");
                Assert.IsTrue(cells.Count > 0, "CKCells should contain at least one cell.");
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

                var cellRef = new DummyRectRef(0, 0, 0, 0);
                CKCells cells = ckTable.Converters.GetCells(ckTable, ckTable, cellRef);

                CKCell firstCell = cells[1];
                Assert.IsNotNull(firstCell, "The indexer should return a valid CKCell for a valid index.");

                Assert.ThrowsException<ArgumentOutOfRangeException>(() => { var _ = cells[0]; });
                Assert.ThrowsException<ArgumentOutOfRangeException>(() => { var _ = cells[cells.Count + 1]; });
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

                var cellRef = new DummyRectRef(0, 0, 0, 0);
                CKCells cells = ckTable.Converters.GetCells(ckTable, ckTable, cellRef);

                int enumeratedCount = cells.Count();
                Assert.AreEqual(cells.Count, enumeratedCount, "Enumeration should yield the same number of cells as the Count property.");
            });
        }
    }
}
