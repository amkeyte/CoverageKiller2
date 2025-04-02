using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.Linq;
using System.Reflection;
using Word = Microsoft.Office.Interop.Word;

namespace CoverageKiller2.DOM.Tables
{
    [TestClass]
    public class CKTableGridTests
    {
        [TestMethod]
        public void CKTableGrid_All_Cells_In_Span_Reference_Same_COMCell()
        {
            LiveWordDocument.WithTestDocument(doc =>
            {
                var table = TestHelpers.FindNthMergedTable(doc, 1);
                var grid = CKTableGrid.GetInstance(table.COMTable);

                var master = grid.GetMasterCells().First();
                var (rowSpan, colSpan) = grid.GetCellSpan(master.COMCell);

                for (int r = 0; r < rowSpan; r++)
                {
                    for (int c = 0; c < colSpan; c++)
                    {
                        int row = master.GridRow + r;
                        int col = master.GridCol + c;

                        var cell = GetGridCell(grid, row, col);
                        Assert.IsNotNull(cell, $"Missing cell at ({row},{col})");
                        Assert.AreEqual(master.COMCell.Range.Start, cell.COMCell.Range.Start,
                            $"Cell at ({row},{col}) has unexpected Range.Start");
                    }
                }
            });
        }
        private static GridCell GetGridCell(CKTableGrid grid, int row, int col)
        {
            var field = typeof(CKTableGrid).GetField("_grid", BindingFlags.NonPublic | BindingFlags.Instance);
            var gridArray = field?.GetValue(grid) as GridCell[,];
            return gridArray?[row, col];
        }
        private const int TestTableNumber = 1;
        [TestMethod]
        public void CKTableGrid_MasterCell_Has_Correct_Span()
        {
            LiveWordDocument.WithTestDocument(doc =>
            {
                var table = TestHelpers.FindNthMergedTable(doc, 1);
                var grid = CKTableGrid.GetInstance(table.COMTable);

                var master = grid.GetMasterCells().First();
                var (rowSpan, colSpan) = grid.GetCellSpan(master.COMCell);

                // Validate that the span is at least 2x2 — adjust as needed for test doc
                Assert.IsTrue(rowSpan >= 1);
                Assert.IsTrue(colSpan >= 1);
            });
        }
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
    }
}
