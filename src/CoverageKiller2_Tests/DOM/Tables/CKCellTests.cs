using CoverageKiller2.Test;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Serilog;
using Word = Microsoft.Office.Interop.Word;

namespace CoverageKiller2.DOM.Tables
{
    [TestClass]
    public class CKCellTests
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
        public void CellRef_HasCorrectRowAndColumnIndex()
        {
            var table = _testFile.Tables[_testTableIndex];
            var cell = table.Cell(1);
            Assert.AreEqual(1, cell.RowIndex);
            Assert.AreEqual(1, cell.ColumnIndex);
            Assert.IsNotNull(cell.CellRef);
            Assert.AreEqual(1, cell.CellRef.RowIndex);
            Assert.AreEqual(1, cell.CellRef.ColumnIndex);
        }

        [TestMethod]
        public void COMCell_IsNotNull_AndReturnsSameTextAsDirectInterop()
        {
            var table = _testFile.Tables[_testTableIndex];
            var cell = table.Cell(1);
            var expected = cell.COMCell.Range.Text.Trim();
            var actual = table.COMTable.Cell(1, 1).Range.Text.Trim();
            Assert.AreEqual(expected, actual, "CKCell should wrap the correct Word.Cell COM object.");
        }

        [TestMethod]
        public void BackgroundColor_SetGet_WorksCorrectly()
        {
            var table = _testFile.Tables[_testTableIndex];
            var cell = table.Cell(1);

            cell.BackgroundColor = Word.WdColor.wdColorGray20;
            Assert.AreEqual(Word.WdColor.wdColorGray20, cell.BackgroundColor);
        }

        [TestMethod]
        public void ForegroundColor_SetGet_WorksCorrectly()
        {
            var table = _testFile.Tables[_testTableIndex];
            var cell = table.Cell(1);

            cell.ForegroundColor = Word.WdColor.wdColorWhite;
            Assert.AreEqual(Word.WdColor.wdColorWhite, cell.ForegroundColor);
        }

        [TestMethod]
        public void Refresh_UpdatesRange()
        {
            var table = _testFile.Tables[_testTableIndex];
            var cell = table.Cell(1);
            cell.IsDirty = true;
            _ = cell.Text; //force refresh
            Assert.IsNotNull(cell.COMRange);
        }
    }
}
