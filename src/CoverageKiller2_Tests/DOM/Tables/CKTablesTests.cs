using CoverageKiller2.Test;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Serilog;
using System.Runtime.InteropServices;
using Word = Microsoft.Office.Interop.Word;

namespace CoverageKiller2.DOM.Tables
{
    [TestClass()]
    public class CKTablesTests
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
        public void Constructor_ShouldInitializeTablesCollection()
        {
            Word.Document comDocument = null;

            try
            {
                var tables = _testFile.Tables;
                comDocument = _testFile.GiveMeCOMDocumentIWillOwnItAndPromiseToCleanUpAfterMyself();

                Assert.IsNotNull(tables, "CKTables should not be null.");
                Assert.AreEqual(comDocument.Tables.Count, tables.Count, "CKTables.Count should match underlying Word.Tables.Count.");
                Assert.IsTrue(tables.Count > 0, "Test document must contain at least one table.");

                var first = tables[1];
                Assert.IsNotNull(first);
                Assert.AreEqual(comDocument.Tables[1].Range.Text, first.COMTable.Range.Text);
            }
            finally
            {
                if (comDocument != null)
                {
                    Marshal.ReleaseComObject(comDocument);
                }
            }
        }


        [TestMethod]
        public void Contains_ShouldReturnTrueForOwnedCell()
        {
            var table = _testFile.Tables[1];
            var cell = table.COMTable.Cell(1, 1);

            Assert.IsTrue(table.Contains(cell), "Expected Contains(cell) to return true for a cell in the table.");
        }

        [TestMethod]
        public void Contains_ShouldReturnFalseForExternalCell()
        {
            if (_testFile.Tables.Count < 2)
                Assert.Inconclusive("Test file must have at least two tables.");

            var table1 = _testFile.Tables[1];
            var table2 = _testFile.Tables[2];
            var foreignCell = table2.COMTable.Cell(1, 1);

            Assert.IsFalse(table1.Contains(foreignCell), "Expected Contains(cell) to return false for a cell in a different table.");
        }

        [TestMethod]
        public void ItemOf_ShouldReturnOwningTableForCell()
        {
            var wordCell = _testFile.Tables[1].COMTable.Cell(1, 1);
            var owningTable = _testFile.Tables.ItemOf(wordCell);

            Assert.IsNotNull(owningTable);
            Assert.IsTrue(RangeSnapshot.FastMatch(
                _testFile.Tables[1].COMTable.Range,
                owningTable.COMTable.Range));
        }


        [TestMethod]
        public void Add_WithRange_ShouldInsertTableAtSpecifiedLocation()
        {
            var insertRange = _testFile.Range();
            var added = _testFile.Tables.Add(insertRange, 2, 2);

            Assert.IsNotNull(added);
            Assert.AreEqual(2, added.COMTable.Rows.Count);
            Assert.AreEqual(2, added.COMTable.Columns.Count);

            // Ensure the inserted table starts at the expected range
            Assert.IsTrue(added.COMRange.Start <= insertRange.Start);
        }
        [TestMethod]
        public void IndexOf_ShouldReturnCorrectIndexForCKTable()
        {
            Log.Verbose("getting Tables from Document");
            var table = _testFile.Tables[1];
            Log.Verbose("checking IndexOf");
            int index = _testFile.Tables.IndexOf(table);

            Assert.AreEqual(0, index, "IndexOf should return 0 for first table.");
        }
    }
}
