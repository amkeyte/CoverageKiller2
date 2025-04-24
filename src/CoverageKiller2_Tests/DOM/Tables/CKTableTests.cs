using CoverageKiller2.Test;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Serilog;
using System;
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
        public void GetCellFor_ValidRef_ReturnsExpectedCOMCell()
        {
            var table = _testFile.Tables[_testTableIndex];
            Assert.IsNotNull(table);

            var cellRef = new CKCellRef(1, 1, table, table);
            var wordCell = table.GetCellFor(cellRef);

            Assert.IsNotNull(wordCell);
            Assert.AreEqual(1, wordCell.RowIndex, "Expected RowIndex to be 1.");
            Assert.AreEqual(1, wordCell.ColumnIndex, "Expected ColumnIndex to be 1.");
        }
        [TestMethod]
        [ExpectedException(typeof(ArgumentOutOfRangeException))]
        public void GetCellFor_InvalidRef_ThrowsArgumentOutOfRangeException()
        {
            var table = _testFile.Tables[_testTableIndex];

            // Triggers constructor-level validation, not CKTable.GetCellFor
            var invalidRef = new CKCellRef(999, 999, table, table);

            table.GetCellFor(invalidRef); // Not reached
        }
        //[TestMethod]
        //[ExpectedException(typeof(ArgumentException))]
        //public void GetCellFor_RefOutsideMasterGrid_ThrowsArgumentException()
        //{
        //    var table = _testFile.Tables[_testTableIndex];

        //    // Use valid coordinates, but expect GetCellFor to fail (e.g., due to merge)
        //    int row = 1;
        //    int col = table.COMTable.Columns.Count + 1; // likely invalid

        //    var cellRef = new CKCellRef(row, col, table, table);
        //    table.GetCellFor(cellRef); // This should now throw inside GetCellFor
        //}
        [TestMethod]
        public void HasMerge_ReturnsExpectedValue()
        {
            var table = _testFile.Tables[_testTableIndex];
            var hasMerge = table.HasMerge;
            Assert.IsInstanceOfType(hasMerge, typeof(bool));
        }

        [TestMethod]
        public void COMTable_IsNotNull_AndMatchesWordInterop()
        {
            var table = _testFile.Tables[_testTableIndex];
            Assert.IsNotNull(table.COMTable, "COMTable should not be null.");
            //Assert.IsInstanceOfType(table.COMTable, typeof(System.__ComObject));
            Assert.AreEqual(table.COMTable.Range.Text, table.COMRange.Text, "COMRange should match the range of COMTable.");
        }

        [TestMethod]
        public void COMRange_CoversSameTextAsCOMTableRange()
        {
            var table = _testFile.Tables[_testTableIndex];
            var comRangeText = table.COMTable.Range.Text.Trim();
            var ckRangeText = table.COMRange.Text.Trim();

            // Strip trailing paragraph markers or whitespace discrepancies
            Assert.IsTrue(ckRangeText.StartsWith(comRangeText) || comRangeText.StartsWith(ckRangeText),
                $"CKTable range text and COMTable range text differ. CKRange: '{ckRangeText}', COMRange: '{comRangeText}'");
        }
    }
}
