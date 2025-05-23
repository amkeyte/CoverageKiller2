﻿using CoverageKiller2.Test;
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
            workSpace.ShowDebuggerWindow();
            var ckTable = workSpace.CloneFrom(_testFile.Tables[_testTableIndex]);
            var COMTable = ckTable.COMTable;
            var crawler = new GridCrawler5(COMTable);

            var clonedTable = crawler.PrepareTable(COMTable);
            Log.Debug($"\n\nTable {_testTableIndex}: {GridCrawler5.FlattenTableText(COMTable.Range.Text)}");

            var textGrid = crawler.ParseTableText(clonedTable);
            Log.Debug(GridCrawler5.DumpGrid(textGrid, nameof(GridCrawler5.ParseTableText)));

            //show the result.
            Log.Debug(GridCrawler5.DumpGrid(textGrid));

            //Assert.AreEqual(wordTable.COMRange.Cells.Count, (long)textGrid.Count);
            //Assert.IsTrue(CKTextHelper.ScrunchEquals(wordTable.COMTable.Cell(1, 1).Range.Text, textGrid[1][1]));
            //Assert.IsTrue(CKTextHelper.ScrunchEquals(wordTable.COMTable.Cell(3, 3).Range.Text, textGrid[3][3]));
        }

        [TestMethod]
        public void Ctor_Test()
        {
            var workSpace = _testFile.Application.GetShadowWorkspace();
            workSpace.ShowDebuggerWindow();
            var ckTable = workSpace.CloneFrom(_testFile.Tables[_testTableIndex]);
            var COMTable = ckTable.COMTable;
            var crawler = new GridCrawler5(COMTable);

            Log.Debug(GridCrawler5.DumpGrid(crawler.Grid));
        }
        [TestMethod]
        public void NormalizebyWidth_Test()
        {
            var workSpace = _testFile.Application.GetShadowWorkspace();
            workSpace.ShowDebuggerWindow();
            var ckTable = workSpace.CloneFrom(_testFile.Tables[_testTableIndex]);
            var COMTable = ckTable.COMTable;

            var crawler = new GridCrawler5(COMTable);
            var clonedTable = crawler.PrepareTable(COMTable);

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
            var workSpace = _testFile.Application.GetShadowWorkspace();
            workSpace.ShowDebuggerWindow();
            var ckTable = workSpace.CloneFrom(_testFile.Tables[_testTableIndex]);
            var COMTable = ckTable.COMTable;
            var crawler = new GridCrawler5(COMTable);

            var clonedTable = crawler.PrepareTable(COMTable);

            var textGrid = crawler.ParseTableText(clonedTable);
            Log.Debug(GridCrawler5.DumpGrid(textGrid, nameof(GridCrawler5.ParseTableText)));

            var masterGrid = crawler.GetMasterGrid(clonedTable);
            Log.Debug(GridCrawler5.DumpGrid(masterGrid, nameof(GridCrawler5.GetMasterGrid)));

            var normalizedGrid = crawler.NormalizeByWidth(masterGrid);
            Log.Debug(GridCrawler5.DumpGrid(normalizedGrid, nameof(GridCrawler5.NormalizeByWidth)));

            var horizontalGrid = crawler.CrawlHoriz(textGrid, normalizedGrid);
            Log.Debug(GridCrawler5.DumpGrid(horizontalGrid, nameof(GridCrawler5.CrawlHoriz)));

            var verticalGrid = crawler.CrawlVertically(0, textGrid, normalizedGrid);
            Log.Debug(GridCrawler5.DumpGrid(normalizedGrid, nameof(GridCrawler5.CrawlVertically)));

            //Assert.Fail();
        }
        [TestMethod]
        public void CrawlHoriz_Test()
        {
            var workSpace = _testFile.Application.GetShadowWorkspace();
            workSpace.ShowDebuggerWindow();
            var ckTable = workSpace.CloneFrom(_testFile.Tables[_testTableIndex]);
            var COMTable = ckTable.COMTable;
            var crawler = new GridCrawler5(COMTable);

            var clonedTable = crawler.PrepareTable(COMTable);

            var textGrid = crawler.ParseTableText(clonedTable);
            Log.Debug(GridCrawler5.DumpGrid(textGrid, nameof(GridCrawler5.ParseTableText)));

            var masterGrid = crawler.GetMasterGrid(clonedTable);
            Log.Debug(GridCrawler5.DumpGrid(masterGrid, nameof(GridCrawler5.GetMasterGrid)));

            var normalizedGrid = crawler.NormalizeByWidth(masterGrid);
            Log.Debug(GridCrawler5.DumpGrid(normalizedGrid, nameof(GridCrawler5.NormalizeByWidth)));

            var horizGrid = crawler.CrawlHoriz(textGrid, normalizedGrid);

            Log.Debug(GridCrawler5.DumpGrid(textGrid, nameof(GridCrawler5.ParseTableText)));
            Log.Debug(GridCrawler5.DumpGrid(normalizedGrid, nameof(GridCrawler5.CrawlHoriz)));


            //Assert.Fail();
        }

        [TestMethod]
        public void AnalyzeTableRecursively_Test()
        {
            var workSpace = _testFile.Application.GetShadowWorkspace();
            workSpace.ShowDebuggerWindow();

            var ckTable = workSpace.CloneFrom(_testFile.Tables[_testTableIndex]);
            var COMTable = ckTable.COMTable;
            var crawler = new GridCrawler5(COMTable);

            Log.Debug($"\n\nAnalyzing Table {_testTableIndex} recursively.");
            var mergedGrid = crawler.AnalyzeTableRecursively(COMTable);

            Log.Debug(GridCrawler5.DumpGrid(mergedGrid, "Merged Crawl Result"));
        }


    }
}
