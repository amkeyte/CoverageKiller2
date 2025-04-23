using CoverageKiller2.Test;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Serilog;

namespace CoverageKiller2.DOM.Tables
{
    /// <summary>
    /// Unit tests for the <see cref="CKTable"/> and related cell access methods.
    /// </summary>
    /// <remarks>
    /// Version: CK2.00.01.0002
    /// </remarks>
    [TestClass]
    public class CKTableTests
    {
        //******* Standard Benchmark Rigging ********
        static int _testTableIndex = 1;
        static int _iterationCount = 1;

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
        public void Constructor_BindsTableCorrectly()
        {
            var table = _testFile.Tables[_testTableIndex];
            Assert.IsNotNull(table);
            Assert.AreEqual(_testFile, table.Document);
        }

        [TestMethod]
        public void Cell_ReturnsExpectedCKCell()
        {
            var table = _testFile.Tables[_testTableIndex];
            var cell = table.Cell(1);
            Assert.IsNotNull(cell);
            var cellRef = new CKCellRef(1, 1, table, table);
            Assert.AreEqual(1, cell.CellRef.ColumnIndex);
            Assert.AreEqual(1, cell.CellRef.RowIndex);
        }

        [TestMethod]
        public void Contains_ValidCell_ReturnsTrue()
        {
            var table = _testFile.Tables[_testTableIndex];
            var wordCell = table.COMTable.Cell(1, 1);
            Assert.IsTrue(table.Contains(wordCell));
        }

        [TestMethod]
        public void IndexOf_ReturnsCorrectIndex()
        {
            var table = _testFile.Tables[_testTableIndex];
            var wordCell = table.COMTable.Cell(1, 1);
            var index = table.IndexOf(wordCell);
            Assert.IsTrue(index > 0);
        }

        [TestMethod]
        public void Columns_CountMatchesCOMTable()
        {
            var table = _testFile.Tables[_testTableIndex];
            var expected = table.COMTable.Columns.Count;
            var actual = table.Columns.Count;
            Assert.AreEqual(expected, actual);
        }

        [TestMethod]
        public void Rows_CountMatchesCOMTable()
        {
            var table = _testFile.Tables[_testTableIndex];
            var expected = table.COMTable.Rows.Count;
            var actual = table.Rows.Count;
            Assert.AreEqual(expected, actual);
        }

        [TestMethod]
        public void HasMerge_ReturnsExpectedValue()
        {
            var table = _testFile.Tables[_testTableIndex];
            var hasMerge = table.HasMerge;
            Assert.IsInstanceOfType(hasMerge, typeof(bool));
        }
    }
}
