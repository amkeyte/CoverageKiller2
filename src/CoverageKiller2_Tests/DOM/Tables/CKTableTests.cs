using CoverageKiller2.Test;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Serilog;

namespace CoverageKiller2.DOM.Tables
{
    /// <summary>
    /// Unit tests for the <see cref="CKTable"/> and related cell access methods.
    /// </summary>
    /// <remarks>
    /// Version: CK2.00.01.0001
    /// </remarks>
    [TestClass]
    public class CKTableTests
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

        [TestMethod]
        public void Constructor_ShouldInitializeCKTable()
        {
            var wordTable = _testFile.Tables[1].COMTable;
            var table = new CKTable(wordTable, _testFile.Tables);

            Assert.IsNotNull(table);
            Assert.IsNotNull(table.COMTable);
            Assert.AreEqual(wordTable.Range.Text, table.COMTable.Range.Text);
        }

        [TestMethod]
        public void Contains_ShouldRecognizeTableCell()
        {
            var table = _testFile.Tables[1];
            var cell = table.COMTable.Cell(1, 1);

            Assert.IsTrue(table.Contains(cell), "Expected Contains(cell) to return true for a table-owned cell.");
        }

        [TestMethod]
        public void Cell_ByIndex_ShouldReturnCorrectCKCell()
        {
            var table = _testFile.Tables[1];
            var ckCell = table.Cell(1); // one-based index

            Assert.IsNotNull(ckCell);
            Assert.AreEqual(1, ckCell.RowIndex);
            Assert.AreEqual(1, ckCell.ColumnIndex);
            Assert.IsTrue(ckCell.COMCell.Range.Text.Length > 0 || ckCell.COMCell.Range.Text == "\r\a"); // empty cell
        }

        [TestMethod]
        public void Cell_ByRef_ShouldReturnMatchingCKCell()
        {
            var table = _testFile.Tables[1];
            var cell = table.COMTable.Cell(1, 1);

            var cellRef = new CKCellRef(
                rowIndex: cell.RowIndex,
                colIndex: cell.ColumnIndex,
                snapshot: new RangeSnapshot(cell.Range),
                parent: table
            );

            var ckCell = table.Cell(cellRef);

            Assert.IsNotNull(ckCell);
            Assert.AreEqual(1, ckCell.RowIndex);
            Assert.AreEqual(1, ckCell.ColumnIndex);
            Assert.IsTrue(ckCell.COMCell.Range.COMEquals(cell.Range));
        }

        [TestMethod]
        public void IndexOf_Cell_ShouldReturnCorrectLinearIndex()
        {
            Log.Debug("First Run");
            var table1 = _testFile.Tables[1];

            var cellRef1 = new CKCellRef(1, 1, table1);
            var first1 = table1.Cell(cellRef1);
            int index1 = table1.IndexOf(first1.COMCell);

            Assert.AreEqual(1, index1, "First cell should have a one-based index of 1.");
            Log.Debug("Second Run");
            var table2 = _testFile.Tables[1];

            var cellRef2 = new CKCellRef(1, 1, table2);
            var first2 = table2.Cell(cellRef2);
            int index2 = table2.IndexOf(first2.COMCell);

            Assert.AreEqual(1, index1, "First cell should have a one-based index of 1.");
        }
    }
}
