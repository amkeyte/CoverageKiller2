using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.Diagnostics;
using System.Linq;

namespace CoverageKiller2.DOM.Tables
{
    [TestClass]
    public class TableGridCrawler3Tests
    {
        public int UseTable = 2;

        [TestMethod]
        public void VisualizeMergedCellsInShadowGrid()
        {
            LiveWordDocument.WithTestDocument(doc =>
            {
                var sourceTable = doc.Tables[UseTable].COMTable;
                var app = doc.Application;

                var shadow = new ShadowWorkspace(app);

                try
                {
                    shadow.ShowDebuggerWindow(keepOpen: true); // Flip to false for CI safety

                    var clone = shadow.CloneTable(sourceTable);
                    var grid = TableGridCrawler3.NormalizeVisualGrid(clone);

                    Assert.IsNotNull(grid, "Grid was null.");
                    Assert.IsTrue(grid.Count > 0, "Grid was empty.");

                    TableGridCrawler3.PrepareGridForLayout(grid);
                    TableGridCrawler3.ColorMasterCells(grid);

                    Debug.WriteLine("=== Grid with Merged Cell Highlighting ===");
                    Debug.WriteLine(TableGridCrawler3.DumpGrid(grid));
                }
                finally
                {
                    shadow.Dispose(); // will skip cleanup if keepOpen=true
                }
            });
        }

        [TestMethod]
        public void NormalizeVisualGrid_BuildsCorrectJaggedGrid()
        {
            LiveWordDocument.WithTestDocument(doc =>
            {
                int tableIndex = UseTable;
                var wordTable = doc.Tables[tableIndex].COMTable;

                var grid = TableGridCrawler3.NormalizeVisualGrid(wordTable);

                Assert.IsNotNull(grid, "Grid is null");
                Assert.IsTrue(grid.Count > 0, "Grid has no rows");
                Assert.IsTrue(grid.All(row => row.Count > 0), "One or more rows are empty");

                Debug.WriteLine("=== DumpGrid ===");
                Debug.WriteLine(TableGridCrawler3.DumpGrid(grid));
            });
        }

        [TestMethod]
        public void ZombieCells_CanHaveNullMasterDuringConstruction()
        {
            LiveWordDocument.WithTestDocument(doc =>
            {
                var wordTable = doc.Tables[UseTable].COMTable;
                var grid = TableGridCrawler3.NormalizeVisualGrid(wordTable);

                var zombies = grid.SelectMany(r => r).Where(c => c.IsDummy).ToList();

                Assert.IsTrue(zombies.Count > 0, "No zombie cells found — are you testing on a merged table?");
                // Do not assert master != null — we allow unassigned initially
                int orphanCount = zombies.Count(z => z.MasterCell == null);
                Debug.WriteLine($"Zombie count: {zombies.Count}, orphaned zombies: {orphanCount}");
            });
        }
        [TestMethod]
        public void MergeSpan_Detection_Works()
        {
            LiveWordDocument.WithTestDocument(doc =>
            {
                var wordTable = doc.Tables[UseTable].COMTable;
                var grid = TableGridCrawler3.NormalizeVisualGrid(wordTable);

                var masters = grid.SelectMany(row => row)
                                  .Where(c => c.IsMasterCell && !c.IsDummy && c.COMCell != null)
                                  .ToList();

                Assert.IsTrue(masters.Count > 0, "No master cells found in grid");

                Debug.WriteLine("=== Merge Span Analysis ===");
                foreach (var master in masters)
                {
                    Debug.WriteLine($"Master [{master.GridRow},{master.GridCol}] " +
                        $"-> RowSpan: {master.DetectedRowSpan}, ColSpan: {master.DetectedColSpan}");

                    Assert.IsTrue(master.DetectedRowSpan >= 1, "Invalid RowSpan");
                    Assert.IsTrue(master.DetectedColSpan >= 1, "Invalid ColSpan");
                }
            });
        }
        [TestMethod]
        public void AllRows_AreSameLength()
        {
            LiveWordDocument.WithTestDocument(doc =>
            {
                var wordTable = doc.Tables[UseTable].COMTable;
                var grid = TableGridCrawler3.NormalizeVisualGrid(wordTable);

                int expectedCols = grid[1].Count;
                Assert.IsTrue(expectedCols > 0, "Grid columns not detected");

                foreach (var row in grid)
                {
                    Assert.AreEqual(expectedCols, row.Count, $"Row mismatch: expected {expectedCols}, got {row.Count}");
                }
            });
        }
    }
}
