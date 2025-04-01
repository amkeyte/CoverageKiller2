using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.Linq;

namespace CoverageKiller2.DOM.Tables
{
    [TestClass]
    public class CKTableTests
    {
        [TestMethod]
        public void FromRange_ShouldReturnValidCKTable()
        {
            LiveWordDocument.WithTestDocument(LiveWordDocument.DefaultTestFile, doc =>
            {
                var range = doc.Tables[1].Range;
                var table = CKTable.FromRange(range);

                Assert.IsNotNull(table);
                Assert.IsNotNull(table.COMTable);
                Assert.AreEqual(range.Text, table.COMTable.Range.Text);
            });
        }

        [TestMethod]
        public void IndexesOf_WordCells_ShouldReturnCorrectIndexes()
        {
            LiveWordDocument.WithTestDocument(LiveWordDocument.DefaultTestFile, doc =>
            {
                var table = CKTable.FromRange(doc.Tables[1].Range);
                var cellList = doc.Tables[1].Range.Cells;
                var indexes = table.IndexesOf(cellList).ToList();

                Assert.AreEqual(cellList.Count, indexes.Count);
                Assert.IsTrue(indexes.All(i => i >= 0));
            });
        }

        [TestMethod]
        public void Cell_FromCKCellRef_ShouldReturnCorrectCell()
        {
            LiveWordDocument.WithTestDocument(LiveWordDocument.DefaultTestFile, doc =>
            {
                var table = CKTable.FromRange(doc.Tables[1].Range);
                var firstCell = doc.Tables[1].Cell(1, 1);
                var refCell = new CKCellRef(firstCell);
                var ckCell = table.Cell(refCell);

                Assert.AreEqual(1, ckCell.WordRow);
                Assert.AreEqual(1, ckCell.WordColumn);
                Assert.AreEqual(firstCell.Range.Text, ckCell.COMCell.Range.Text);
            });
        }
        [TestMethod]
        public void Constructor_ShouldInitializeCKTable()
        {
            LiveWordDocument.WithTestDocument(LiveWordDocument.DefaultTestFile, doc =>
            {
                var wordTable = doc.Tables[1];
                var table = new CKTable(wordTable);

                Assert.IsNotNull(table);
                Assert.IsNotNull(table.COMTable);
                Assert.AreEqual(wordTable.Range.Text, table.COMTable.Range.Text);
            });
        }
        [TestMethod]
        public void IndexesOf_CKCells_ShouldMatchMasterCells()
        {
            LiveWordDocument.WithTestDocument(LiveWordDocument.DefaultTestFile, doc =>
            {
                var table = CKTable.FromRange(doc.Tables[1].Range);
                var firstCell = doc.Tables[1].Cell(1, 1);
                var refCell = new CKCellRef(firstCell);
                var ckCell = table.Cell(refCell);
                var ckCells = CKCells.FromRef(table, ckCell.CellRef);

                var indexes = table.IndexesOf(ckCells).ToList();

                Assert.AreEqual(1, indexes.Count);
                Assert.IsTrue(indexes[0] >= 0);
            });
        }
    }
}
