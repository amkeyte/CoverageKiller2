using CoverageKiller2._TestOperators;
using CoverageKiller2.Test;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Serilog;
using System.Linq;

namespace CoverageKiller2.DOM.Tables
{
    /// <summary>
    /// Unit tests for <see cref="CKTableGrid"/> behavior and integrity.
    /// </summary>
    /// <remarks>
    /// Version: CK2.00.00.0002
    /// </remarks>
    [TestClass]
    public class CKTableGridTests
    {
        //******* Standard Rigging ********
        public TestContext TestContext { get; set; }
        private string _testFilePath;
        private CKDocument _testFile;

        [TestInitialize]
        public void Setup()
        {
            Log.Information($"Running test => {GetType().Name}::{TestContext.TestName}");
            _testFilePath = RandomTestHarness.TestFile1;
            _testFile = RandomTestHarness.GetTempDocumentFrom(_testFilePath);
        }
        [TestCleanup]
        public void Cleanup()
        {
            RandomTestHarness.CleanUp(_testFile, force: true);
            Log.Information($"Completed test => {GetType().Name}::{TestContext.TestName}; status: {TestContext.CurrentTestOutcome}");
        }
        //******* End Standard Rigging ********

        private CKDocument _doc;
        private const int TestTableNumber = 1;

        [TestInitialize]
        public void SetUp()
        {
            _doc = RandomTestHarness.GetTempDocumentFrom(RandomTestHarness.TestFile1);
        }

        [TestCleanup]
        public void TearDown()
        {
            RandomTestHarness.CleanUp(_doc);
            _doc = null;
        }

        [TestMethod]
        public void CKTableGrid_All_Cells_In_Span_Reference_Same_COMCell()
        {
            var table = TestHelpers.FindNthMergedTable(_doc, 1);
            var grid = CKTableGrid.GetInstance(table, table.COMTable);

            var master = grid.GetMasterCells().First();
            var (rowSpan, colSpan) = grid.GetCellSpan(master.COMCell);

            for (int r = 0; r < rowSpan; r++)
            {
                for (int c = 0; c < colSpan; c++)
                {
                    int row = master.GridRow + r;
                    int col = master.GridCol + c;

                    var cell = grid._grid[row, col];
                    Assert.IsNotNull(cell, $"Missing cell at ({row},{col})");
                    Assert.AreEqual(master.COMCell.Range.Start, cell.COMCell.Range.Start,
                        $"Cell at ({row},{col}) has unexpected Range.Start");
                }
            }
        }

        [TestMethod]
        public void CKTableGrid_MasterCell_Has_Correct_Span()
        {
            var table = TestHelpers.FindNthMergedTable(_doc, 1);
            var grid = CKTableGrid.GetInstance(table, table.COMTable);

            var master = grid.GetMasterCells().First();
            var (rowSpan, colSpan) = grid.GetCellSpan(master.COMCell);

            Assert.IsTrue(rowSpan >= 1);
            Assert.IsTrue(colSpan >= 1);
        }

        [TestMethod]
        public void CKTableGrid_GetInstance_ReturnsSameInstanceForSameTableRange()
        {
            Assert.IsTrue(_doc.Tables.Count >= TestTableNumber);
            CKTable wordTable = _doc.Tables[TestTableNumber];

            CKTableGrid grid1 = CKTableGrid.GetInstance(wordTable, wordTable.COMTable);
            CKTableGrid grid2 = CKTableGrid.GetInstance(wordTable, wordTable.COMTable);

            Assert.IsNotNull(grid1);
            Assert.AreSame(grid1, grid2);
        }

        [TestMethod]
        public void CKTableGrid_GetCellAt_ValidCoordinates_ReturnsGridCell()
        {
            Assert.IsTrue(_doc.Tables.Count >= TestTableNumber);
            CKTable wordTable = _doc.Tables[TestTableNumber];

            CKTableGrid grid = CKTableGrid.GetInstance(wordTable, wordTable.COMTable);

            Assert.IsTrue(grid.RowCount > 0);
            Assert.IsTrue(grid.ColCount > 0);

            var cell = grid.GetMasterCells(new CKGridCellRef(0, 0, 0, 0)).FirstOrDefault();
            Assert.IsNotNull(cell);
        }

        [TestMethod]
        public void CKTableGrid_GetCellAt_InvalidCoordinates_ReturnsNull()
        {
            Assert.IsTrue(_doc.Tables.Count >= TestTableNumber);
            CKTable wordTable = _doc.Tables[TestTableNumber];

            CKTableGrid grid = CKTableGrid.GetInstance(wordTable, wordTable.COMTable);

            var cell = grid.GetMasterCells(new CKGridCellRef(999, 999, 999, 999)).FirstOrDefault();
            Assert.IsNull(cell);
        }
    }
}
