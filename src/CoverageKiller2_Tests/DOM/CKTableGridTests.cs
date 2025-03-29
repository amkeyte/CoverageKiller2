using CoverageKiller2.Tests;  // Contains LiveWordDocument helper.
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.Linq;
using Word = Microsoft.Office.Interop.Word;

namespace CoverageKiller2.DOM
{
    [TestClass]
    public class CKTableGridTests
    {
        private const int TestTableNumber = 1;

        [TestMethod]
        public void CKTableGrid_GetInstance_ReturnsSameInstanceForSameTableRange()
        {
            LiveWordDocument.WithTestDocument(LiveWordDocument.DefaultTestFile, doc =>
            {
                Assert.IsTrue(doc.Tables.Count >= TestTableNumber);
                Word.Table wordTable = doc.Tables[TestTableNumber];

                CKTableGrid grid1 = CKTableGrid.GetInstance(wordTable);
                CKTableGrid grid2 = CKTableGrid.GetInstance(wordTable);

                Assert.IsNotNull(grid1);
                Assert.AreSame(grid1, grid2);
            });
        }

        [TestMethod]
        public void CKTableGrid_GetCellAt_ValidCoordinates_ReturnsGridCell()
        {
            LiveWordDocument.WithTestDocument(LiveWordDocument.DefaultTestFile, doc =>
            {
                Assert.IsTrue(doc.Tables.Count >= TestTableNumber);
                Word.Table wordTable = doc.Tables[TestTableNumber];

                CKTableGrid grid = CKTableGrid.GetInstance(wordTable);

                Assert.IsTrue(grid.RowCount > 0);
                Assert.IsTrue(grid.ColCount > 0);

                GridCell cell = grid.GetMasterCells(new CKGridCellRef(0, 0, 0, 0)).FirstOrDefault();
                Assert.IsNotNull(cell);
            });
        }

        [TestMethod]
        public void CKTableGrid_GetCellAt_InvalidCoordinates_ReturnsNull()
        {
            LiveWordDocument.WithTestDocument(LiveWordDocument.DefaultTestFile, doc =>
            {
                Assert.IsTrue(doc.Tables.Count >= TestTableNumber);
                Word.Table wordTable = doc.Tables[TestTableNumber];

                CKTableGrid grid = CKTableGrid.GetInstance(wordTable);

                GridCell cell = grid.GetMasterCells(new CKGridCellRef(999, 999, 999, 999)).FirstOrDefault();
                Assert.IsNull(cell);
            });
        }

        //[TestMethod]
        //public void CKTableGrid_Refresh_DoesNotThrowAndReturnsDiffs()
        //{
        //    LiveWordDocument.WithTestDocument(LiveWordDocument.Default, doc =>
        //    {
        //        Assert.IsTrue(doc.Tables.Count >= TestTableNumber);
        //        Word.Table wordTable = doc.Tables[TestTableNumber];
        //        CKTable ckTable = new CKTable(wordTable);

        //        var diffs = CKTableGrid.Refresh(ckTable, true);
        //        Assert.IsNotNull(diffs);
        //    });
        //}
    }
}
