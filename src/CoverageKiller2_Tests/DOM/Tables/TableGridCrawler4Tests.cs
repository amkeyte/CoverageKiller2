using CoverageKiller2.Logging;
using CoverageKiller2.Test;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Serilog;
using System;
using System.Linq;
using Word = Microsoft.Office.Interop.Word;

namespace CoverageKiller2.DOM.Tables
{
    /// <summary>
    /// Tests for the TableGridCrawler4 layout utilities.
    /// </summary>
    /// <remarks>
    /// Version: CK2.00.01.0028
    /// </remarks>
    [TestClass]
    public class TableGridCrawler4Tests
    {
        static int _testTableIndex = 4;

        //******* Standard Rigging ********
        private string _testFilePath;
        private CKDocument _testFile;
        private Tracer Tracer = new Tracer(typeof(TableGridCrawler4Tests));
        [TestInitialize]
        public void Setup()
        {
            _testFilePath = RandomTestHarness.TestFile1;
            _testFile = RandomTestHarness.GetTempDocumentFrom(_testFilePath);
        }

        [TestCleanup]
        public void Cleanup()
        {
            RandomTestHarness.CleanUp(_testFile, force: !_testFile.KeepAlive);
        }
        //******* End Standard Rigging ********

        [TestMethod]
        public void LabelCellsWithCoordinates_ShouldOverwriteAllMasterCells()
        {
            // Arrange
            var ckTable = _testFile.Tables[_testTableIndex];

            // Act
            var clonedTable = TableGridCrawler4.CloneAndPrepareTableLayout(ckTable);

            // Assert
            foreach (Word.Cell cell in clonedTable.COMRange.Cells)
            {
                var expected = $"[{cell.RowIndex},{cell.ColumnIndex}]";
                Assert.IsTrue(CKTextHelper.ScrunchEquals(expected, cell.Range.Text), $"Mismatch at {expected}");
            }
        }

        [TestMethod]
        public void GetMasterGrid_FromCKTable_ReturnsNonEmptyGrid()
        {
            // Arrange
            var ckTable = _testFile.Tables[_testTableIndex];

            // Act
            var grid = TableGridCrawler4.GetMasterGrid(ckTable);

            // Assert
            Assert.IsTrue(grid.Count > 0, "Grid should have rows.");
            Assert.IsTrue(grid.All(row => row.Count > 0), "Each row should have cells.");
        }

        [TestMethod]
        public void DumpGrid_ShouldReturnFormattedString()
        {
            // Arrange
            var ckTable = _testFile.Tables[_testTableIndex];
            var grid = TableGridCrawler4.GetMasterGrid(ckTable);

            // Act
            var output = TableGridCrawler4.DumpGrid(grid);

            // Assert
            Assert.IsFalse(string.IsNullOrWhiteSpace(output));
            Assert.IsTrue(output.Contains("["));
        }

        [TestMethod]
        public void InsertGridAsTableAtEnd_ShouldCreateCorrectDimensionsAndContent()
        {
            //                //table 1 "A\r\aB\r\aC\r\a\r\aD\r\aE\r\aF\r\a\r\aG\r\aH\r\aI\r\a\r\aJ\r\aK\r\aL\r\a\r\a"
            //                //table 2 [w=20]\r\n\a[w=10]\r\n\a\r\n\a\r\n\a[w=10]\r\n\a\r\n\a[w=10]\r\n\a[w=10]\r\n\a[w=10]\r\n\a\r\n\a[w=10]\r\n\a[w=10]\r\n\a[w=10]\r\n\a\r\n\a"
            //                //table 2 ABDE\r\n\aC\r\n\a\r\n\a\r\n\aF\r\n\a\r\n\aG\r\n\aH\r\n\aI\r\n\a\r\n\aJ\r\n\aK\r\n\aL\r\n\a\r\n\a"
            //                //table 3 [1,1]\r\a[1,2]\r\a[1,3]\r\a\r\a[2,1]\r\a\r\a[3,1]\r\a[3,2]\r\a[3,3]\r\a\r\a"
            //                //table 4 ADG\r\aB\r\aC\r\a\r\a\r\aE\r\aF\r\a\r\a\r\aH\r\aI\r\a\r\a"
            //                //xx = "ADG\r\a\r\a\r\a\r\a\r\a\r\a\r\a\r\a\r\aH\r\aI\r\a\r\a"
            //		table 4:  [1,1]\r\a[1,2]\r\a[1,3]\r\a\r\a\r\a[2,2]\r\a\r\a\r\a\r\a[3,2]\r\a\r\a\r\a



            // Arrange
            //var grid = new Base1JaggedList<GridCell2>();
            //for (int r = 1; r <= 2; r++)
            //{
            //    var row = new Base1List<GridCell2>();
            //    for (int c = 1; c <= 3; c++)
            //    {
            //        row.Add(new GridCell2(null, r, c, true));
            //    }
            //    grid.Add(row);
            //}

            //put 1st clone on page.
            var table = TableGridCrawler4.CloneAndPrepareTableLayout(_testFile.Tables[_testTableIndex]);
            //we need a master grid... probably get one back from prev step.
            var grid = TableGridCrawler4.GetMasterGrid(table);
            Tracer.Log($"Dump Grid {nameof(grid)}: \n{TableGridCrawler4.DumpGrid(grid)}\n");


            // drop in zombie cells for wide cells.
            var grid1 = TableGridCrawler4.NormalizeGridByMeasuredWidth(table);
            Tracer.Log($"Dump Grid {nameof(grid1)}: \n{TableGridCrawler4.DumpGrid(grid1)}\n");

            var newTable = TableGridCrawler4.InsertGridAsTableAtEnd(table.Document, grid1);
            Tracer.Log($"\nTable text dump: \n{table.ParsedDebugText[1].DebugDump}");
            // Assert
            Assert.AreEqual(2, newTable.COMTable.Rows.Count, "Inserted row count mismatch.");
            Assert.AreEqual(3, newTable.COMTable.Columns.Count, "Inserted column count mismatch.");

            for (int r = 1; r <= 2; r++)
            {
                for (int c = 1; c <= 3; c++)
                {
                    var expected = $"[{r},{c}]";
                    var actual = newTable.COMTable.Cell(r, c).Range.Text.TrimEnd('\r', '\a');
                    Assert.AreEqual(expected, actual, $"Mismatch at cell [{r},{c}].");
                }
            }
            table.Document.Visible = true;
        }

        [TestMethod]
        [ExpectedException(typeof(ArgumentException))]
        public void InsertGridAsTableAtEnd_ShouldThrowIfGridEmpty()
        {
            var workspace = _testFile.Application.GetShadowWorkspace();
            var emptyGrid = new Base1JaggedList<GridCell2>();
            TableGridCrawler4.InsertGridAsTableAtEnd(workspace, emptyGrid);
            workspace.ShowDebuggerWindow();
            Log.Debug(TableGridCrawler4.DumpGrid(emptyGrid));
        }

        [TestMethod]
        [ExpectedException(typeof(ArgumentNullException))]
        public void InsertGridAsTableAtEnd_ShouldThrowIfWorkspaceNull()
        {
            var dummyGrid = new Base1JaggedList<GridCell2>();
            var row = new Base1List<GridCell2>();
            row.Add(new GridCell2(null, 1, 1, true));
            dummyGrid.Add(row);

            TableGridCrawler4.InsertGridAsTableAtEnd(null, dummyGrid);

        }
    }
}
