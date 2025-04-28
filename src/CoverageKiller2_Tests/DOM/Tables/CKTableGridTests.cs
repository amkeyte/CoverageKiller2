using CoverageKiller2.Test;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Serilog;

namespace CoverageKiller2.DOM.Tables
{
    /// <summary>
    /// Unit tests for <see cref="CKTableGrid"/> behavior and integrity.
    /// </summary>
    /// <remarks>
    /// Version: CK2.00.00.0003
    /// </remarks>
    [TestClass]
    public class CKTableGridTests
    {
        //******* Standard Rigging ********
        public TestContext TestContext { get; set; }
        private string _testFilePath;
        private CKDocument _testFile;
        private const int TestTableNumber = 1;

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

        [TestMethod]
        public void CKTableGrid_GetInstance_ReturnsSameInstanceForSameTableRange()
        {
            Assert.IsTrue(_testFile.Tables.Count >= TestTableNumber);
            CKTable wordTable = _testFile.Tables[TestTableNumber];

            CKTableGrid grid1 = CKTableGrid.GetInstance(wordTable, wordTable.COMTable);
            CKTableGrid grid2 = CKTableGrid.GetInstance(wordTable, wordTable.COMTable);

            Assert.IsNotNull(grid1);
            Assert.AreSame(grid1, grid2);
        }

        [TestMethod]
        public void CKTableGrid_GetMasterCell_ValidCoordinates_ReturnsGridCell()
        {
            Assert.IsTrue(_testFile.Tables.Count >= TestTableNumber);
            CKTable wordTable = _testFile.Tables[TestTableNumber];

            CKTableGrid grid = CKTableGrid.GetInstance(wordTable, wordTable.COMTable);

            Assert.IsTrue(grid.RowCount > 0);
            Assert.IsTrue(grid.ColCount > 0);

            var cellRef = new CKGridCellRef(1, 1, 1, 1); // corrected: 1-based
            var masterCell = grid.GetMasterCell(cellRef);

            Assert.IsNotNull(masterCell);
            Assert.IsTrue(masterCell.IsMasterCell);
        }

        [TestMethod]
        public void CKTableGrid_GetMasterCell_InvalidCoordinates_ReturnsNull()
        {
            Assert.IsTrue(_testFile.Tables.Count >= TestTableNumber);
            CKTable wordTable = _testFile.Tables[TestTableNumber];

            CKTableGrid grid = CKTableGrid.GetInstance(wordTable, wordTable.COMTable);

            var invalidRef = new CKGridCellRef(999, 999, 999, 999); // way out of bounds

            var masterCell = grid.GetMasterCell(invalidRef);

            Assert.IsNull(masterCell, "Expected null when accessing invalid coordinates.");
        }
    }
}
