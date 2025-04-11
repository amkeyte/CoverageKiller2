using CoverageKiller2.Test;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Serilog;
using System.Diagnostics;
using System.Linq;

namespace CoverageKiller2.DOM.Tables
{
    [TestClass]
    public class TableGridCrawler3Tests
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
        public void VisualizeMergedCellsInShadowGrid()
        {

            var sourceTable = _testFile.Tables[2];
            var shadow = _testFile.Application.GetShadowWorkspace(true);
            shadow.ShowDebuggerWindow();


            try
            {
                shadow.ShowDebuggerWindow();
                var cloneTable = shadow.CloneFrom(sourceTable);
                var grid = TableGridCrawler3.NormalizeVisualGrid(cloneTable.COMTable);

                Assert.IsNotNull(grid, "Grid was null.");
                Assert.IsTrue(grid.Count > 0, "Grid was empty.");

                TableGridCrawler3.PrepareGridForLayout(grid);
                TableGridCrawler3.ColorMasterCells(grid);

                TestContext.WriteLine("=== Grid with Merged Cell Highlighting ===");
                TestContext.WriteLine(TableGridCrawler3.DumpGrid(grid));
            }
            finally
            {
                shadow.Dispose();
            }
        }

        [TestMethod]
        public void NormalizeVisualGrid_BuildsCorrectJaggedGrid()
        {
            var wordTable = _testFile.Tables[2].COMTable;
            var grid = TableGridCrawler3.NormalizeVisualGrid(wordTable);

            Assert.IsNotNull(grid, "Grid is null");
            Assert.IsTrue(grid.Count > 0, "Grid has no rows");
            Assert.IsTrue(grid.All(row => row.Count > 0), "One or more rows are empty");

            TestContext.WriteLine("=== DumpGrid ===");
            TestContext.WriteLine(TableGridCrawler3.DumpGrid(grid));
        }

        [TestMethod]
        public void ZombieCells_CanHaveNullMasterDuringConstruction()
        {
            var wordTable = _testFile.Tables[2].COMTable;
            var grid = TableGridCrawler3.NormalizeVisualGrid(wordTable);

            var zombies = grid.SelectMany(r => r).Where(c => c.IsDummy).ToList();

            Assert.IsTrue(zombies.Count > 0, "No zombie cells found — are you testing on a merged table?");
            int orphanCount = zombies.Count(z => z.MasterCell == null);
            Debug.WriteLine($"Zombie count: {zombies.Count}, orphaned zombies: {orphanCount}");
        }

        [TestMethod]
        public void MergeSpan_Detection_Works()
        {
            var wordTable = _testFile.Tables[2].COMTable;
            var grid = TableGridCrawler3.NormalizeVisualGrid(wordTable);

            var masters = grid.SelectMany(row => row)
                              .Where(c => c.IsMasterCell && !c.IsDummy && c.COMCell != null)
                              .ToList();

            Assert.IsTrue(masters.Count > 0, "No master cells found in grid");

            Debug.WriteLine("=== Merge Span Analysis ===");
            foreach (var master in masters)
            {
                Debug.WriteLine($"Master [{master.GridRow},{master.GridCol}] -> RowSpan: {master.DetectedRowSpan}, ColSpan: {master.DetectedColSpan}");
                Assert.IsTrue(master.DetectedRowSpan >= 1, "Invalid RowSpan");
                Assert.IsTrue(master.DetectedColSpan >= 1, "Invalid ColSpan");
            }
        }

        [TestMethod]
        public void AllRows_AreSameLength()
        {
            var wordTable = _testFile.Tables[2].COMTable;
            var grid = TableGridCrawler3.NormalizeVisualGrid(wordTable);

            int expectedCols = grid[1].Count;
            Assert.IsTrue(expectedCols > 0, "Grid columns not detected");

            foreach (var row in grid)
            {
                Assert.AreEqual(expectedCols, row.Count, $"Row mismatch: expected {expectedCols}, got {row.Count}");
            }
        }
    }
}
