using CoverageKiller2.Test;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Serilog;
using System.Linq;

namespace CoverageKiller2.DOM.Tables
{
    /// <summary>
    /// Unit tests for the <see cref="CKTable"/> and related cell access methods.
    /// </summary>
    /// <remarks>
    /// Version: CK2.00.00.0000
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
        public void IndexesOf_WordCells_ShouldReturnCorrectIndexes()
        {
            var table = _testFile.Tables[1];
            var cellList = _testFile.Tables[1].Cells;
            var indexes = table.IndexesOf(cellList).ToList();

            Assert.AreEqual(cellList.Count, indexes.Count);
            Assert.IsTrue(indexes.All(i => i >= 1));
        }



        [TestMethod]
        public void Constructor_ShouldInitializeCKTable()
        {
            var wordTable = _testFile.Tables[1].COMTable;
            var table = new CKTable(wordTable, _testFile.Tables);

            Assert.IsNotNull(table);
            Assert.IsNotNull(table.COMTable);
            Assert.AreEqual(wordTable.Range.Text, table.COMTable.Range.Text);
        }

        //[TestMethod]
        //public void IndexesOf_CKCells_ShouldMatchMasterCells()
        //{
        //    var doc = RandomTestHarness.GetDocument(_testFilePath);
        //    var table = CKTable.FromRange(doc.Tables[1].COMRange, doc);
        //    var firstCell = doc.Tables[1].Cell(1, 1);
        //    var refCell = new CKCellRef(firstCell, table);
        //    var ckCell = table.Cell(refCell);
        //    var ckCells = CKCells.FromRef(table, ckCell.CellRef);

        //    var indexes = table.IndexesOf(ckCells).ToList();

        //    Assert.AreEqual(1, indexes.Count);
        //    Assert.IsTrue(indexes[0] >= 1);
        //}
    }
}
