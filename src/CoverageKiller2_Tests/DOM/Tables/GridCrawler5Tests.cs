using CoverageKiller2.Test;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Serilog;

namespace CoverageKiller2.DOM.Tables
{
    /// <summary>
    /// Stepwise unit tests for GridCrawler5 merge logic.
    /// </summary>
    /// <remarks>
    /// Version: CK2.00.01.0036
    /// </remarks>
    [TestClass]
    public class GridCrawler5StepTests
    {
        //******* Standard Rigging ********
        static int _testTableIndex = 16;

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
        public void ParseTableText_Test()
        {
            var workSpace = _testFile.Application.GetShadowWorkspace();
            var wordTable = _testFile.Tables[_testTableIndex];
            workSpace.ShowDebuggerWindow();

            var crawler = new GridCrawler5(wordTable);
            var clonedTable = crawler.CloneAndPrepareTableLayout(wordTable, workSpace);
            Log.Debug($"\n\nTable {_testTableIndex}: {GridCrawler5.FlattenTableText(clonedTable.RawText)}");

            var textGrid = crawler.ParseTableText(clonedTable);
            Log.Debug(GridCrawler5.DumpGrid(textGrid, nameof(GridCrawler5.ParseTableText)));

            //show the result.
            Log.Debug(GridCrawler5.DumpGrid(textGrid));

            //Assert.AreEqual(wordTable.COMRange.Cells.Count, (long)textGrid.Count);
            //Assert.IsTrue(CKTextHelper.ScrunchEquals(wordTable.COMTable.Cell(1, 1).Range.Text, textGrid[1][1]));
            //Assert.IsTrue(CKTextHelper.ScrunchEquals(wordTable.COMTable.Cell(3, 3).Range.Text, textGrid[3][3]));
        }

        [TestMethod]
        public void NormalizebyWidth_Test()
        {
            var workspace = _testFile.Application.GetShadowWorkspace();
            var wordTable = _testFile.Tables[_testTableIndex];
            workspace.ShowDebuggerWindow();

            var crawler = new GridCrawler5(wordTable);
            var clonedTable = crawler.CloneAndPrepareTableLayout(wordTable, workspace);

            var textGrid = crawler.ParseTableText(clonedTable);
            Log.Debug(GridCrawler5.DumpGrid(textGrid, nameof(GridCrawler5.ParseTableText)));

            var masterGrid = crawler.GetMasterGrid(clonedTable);
            Log.Debug(GridCrawler5.DumpGrid(masterGrid, nameof(GridCrawler5.GetMasterGrid)));

            var normalizedGrid = crawler.NormalizeByWidth(masterGrid);
            Log.Debug(GridCrawler5.DumpGrid(normalizedGrid, nameof(GridCrawler5.NormalizeByWidth)));


            //Assert.Fail();

        }

        [TestMethod]
        public void CrawlVertically_Test()
        {
            var workspace = _testFile.Application.GetShadowWorkspace();
            workspace.ShowDebuggerWindow();
            var wordTable = _testFile.Tables[_testTableIndex];

            var crawler = new GridCrawler5(wordTable);
            var clonedTable = crawler.CloneAndPrepareTableLayout(wordTable, workspace);

            var textGrid = crawler.ParseTableText(clonedTable);
            Log.Debug(GridCrawler5.DumpGrid(textGrid, nameof(GridCrawler5.ParseTableText)));

            var masterGrid = crawler.GetMasterGrid(clonedTable);
            Log.Debug(GridCrawler5.DumpGrid(masterGrid, nameof(GridCrawler5.GetMasterGrid)));

            var normalizedGrid = crawler.NormalizeByWidth(masterGrid);
            Log.Debug(GridCrawler5.DumpGrid(normalizedGrid, nameof(GridCrawler5.NormalizeByWidth)));

            var horizontalGrid = crawler.CrawlHoriz(textGrid, normalizedGrid);
            Log.Debug(GridCrawler5.DumpGrid(horizontalGrid, nameof(GridCrawler5.CrawlHoriz)));

            var verticalGrid = crawler.CrawlVertically(textGrid, normalizedGrid);
            Log.Debug(GridCrawler5.DumpGrid(normalizedGrid, nameof(GridCrawler5.CrawlVertically)));

            //Assert.Fail();
        }
        [TestMethod]
        public void CrawlHoriz_Test()
        {
            var workspace = _testFile.Application.GetShadowWorkspace();
            var wordTable = _testFile.Tables[_testTableIndex];

            var crawler = new GridCrawler5(wordTable);
            var clonedTable = crawler.CloneAndPrepareTableLayout(wordTable, workspace);

            var textGrid = crawler.ParseTableText(clonedTable);
            Log.Debug(GridCrawler5.DumpGrid(textGrid, nameof(GridCrawler5.ParseTableText)));

            var masterGrid = crawler.GetMasterGrid(clonedTable);
            Log.Debug(GridCrawler5.DumpGrid(masterGrid, nameof(GridCrawler5.GetMasterGrid)));

            var normalizedGrid = crawler.NormalizeByWidth(masterGrid);
            Log.Debug(GridCrawler5.DumpGrid(normalizedGrid, nameof(GridCrawler5.NormalizeByWidth)));

            var horizGrid = crawler.CrawlHoriz(textGrid, normalizedGrid);

            Log.Debug(GridCrawler5.DumpGrid(textGrid, nameof(GridCrawler5.ParseTableText)));
            Log.Debug(GridCrawler5.DumpGrid(normalizedGrid, nameof(GridCrawler5.CrawlHoriz)));

            workspace.ShowDebuggerWindow();
            //Assert.Fail();
        }
    }
}
