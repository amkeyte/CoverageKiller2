using CoverageKiller2.Tests;  // Contains LiveWordDocument helper.
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.Collections.Generic;
using Word = Microsoft.Office.Interop.Word;

namespace CoverageKiller2.DOM
{
    [TestClass]
    public class CKTableTests
    {
        [TestMethod]
        public void CKTable_Constructor_LoadsTableSuccessfully()
        {
            LiveWordDocument.WithTestDocument(LiveWordDocument.Default, doc =>
            {
                Assert.IsTrue(doc.Tables.Count > 0);
                Word.Table wordTable = doc.Tables[1];
                CKTable ckTable = new CKTable(wordTable);
                Assert.IsNotNull(ckTable.COMTable);
                Assert.AreEqual(wordTable.Range.Start, ckTable.COMTable.Range.Start);
            });
        }

        [TestMethod]
        public void CKTable_Cell_MatchesDirectWordCell()
        {
            LiveWordDocument.WithTestDocument(LiveWordDocument.Default, doc =>
            {
                Assert.IsTrue(doc.Tables.Count > 0);
                Word.Table wordTable = doc.Tables[1];
                CKTable ckTable = new CKTable(wordTable);

                var cellRef = new CellRefCoord(0, 0, 1);
                CKCell cellFromTable = ckTable.Converters.GetCell(ckTable, cellRef);

                Word.Cell wordCell = wordTable.Cell(1, 1);
                CKCell directCell = new CKCell(ckTable, ckTable, wordCell, 1, 1);

                Assert.AreEqual(directCell.COMCell, cellFromTable.COMCell);
            });
        }

        private const int TestTableNumber = 1;

        [TestMethod]
        public void CKTable_FactoryMethods_SingleCellInTestTable()
        {
            LiveWordDocument.WithTestDocument(LiveWordDocument.Default, doc =>
            {
                Assert.IsTrue(doc.Tables.Count >= TestTableNumber);
                Word.Table wordTable = doc.Tables[TestTableNumber];
                CKTable ckTable = new CKTable(wordTable);

                var cellRef = new CellRefCoord(0, 0, 1);
                CKCell cell = ckTable.Converters.GetCell(ckTable, cellRef);

                Assert.IsNotNull(cell);
            });
        }

        [TestMethod]
        public void CKTable_FactoryMethods_RectangularReferenceInTestTable()
        {
            LiveWordDocument.WithTestDocument(LiveWordDocument.Default, doc =>
            {
                Assert.IsTrue(doc.Tables.Count >= TestTableNumber);
                Word.Table wordTable = doc.Tables[TestTableNumber];
                CKTable ckTable = new CKTable(wordTable);

                var rectRef = new DummyRectRef(0, 0, 1, 1);
                CKCells cells = ckTable.Converters.GetCells(ckTable, ckTable, rectRef);

                Assert.IsNotNull(cells);
                Assert.IsTrue(cells.Count > 0);
            });
        }

        private class DummyRectRef : ICellRef<CKCellsRect>
        {
            public IEnumerable<int> WordCells => new[] { 1, 2, 3, 4 };
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
    }
}
