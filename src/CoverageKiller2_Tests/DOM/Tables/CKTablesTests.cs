using CoverageKiller2.Test;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Serilog;

namespace CoverageKiller2.DOM.Tables
{
    [TestClass()]
    public class CKTablesTests
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
        public void Indexer_ReturnsExpectedTable()
        {
            var tables = _testFile.Tables;
            var table = tables[_testTableIndex];
            Assert.IsNotNull(table);
            Assert.AreEqual(_testFile, table.Document);
        }

        [TestMethod]
        public void Count_MatchesCOMTableCount()
        {
            var tables = _testFile.Tables;
            int expected = _testFile.GiveMeCOMDocumentIWillOwnItAndPromiseToCleanUpAfterMyself()
                .Tables.Count;
            Assert.AreEqual(expected, tables.Count);
        }

        [TestMethod]
        public void IndexOf_ReturnsCorrectIndex()
        {
            var tables = _testFile.Tables;
            var table = tables[_testTableIndex];
            int index = tables.IndexOf(table);
            Assert.AreEqual(_testTableIndex, index);
        }

        [TestMethod]
        public void ItemOf_ReturnsOwningTable()
        {
            var tables = _testFile.Tables;
            var wordCell = tables[_testTableIndex].COMTable.Cell(1, 1);
            var owning = tables.ItemOf(wordCell);
            Assert.IsNotNull(owning);
            Assert.AreEqual(tables[_testTableIndex], owning);
        }

        [TestMethod]
        public void Add_CreatesNewTable()
        {
            var insertAt = _testFile.Range().CollapseToEnd();
            var initialCount = _testFile.Tables.Count;
            var newTable = _testFile.Tables.Add(insertAt, 2, 2);

            Assert.IsNotNull(newTable);
            Assert.AreEqual(initialCount + 1, _testFile.Tables.Count);
            Assert.AreEqual(2, newTable.Rows.Count);
            Assert.AreEqual(2, newTable.Columns.Count);
        }
    }
}
