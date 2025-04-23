using CoverageKiller2.Test;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Serilog;
using System;

namespace CoverageKiller2.DOM.Tables
{
    [TestClass()]
    public class CKColumnTests
    {
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

        private void RunOnEachTestTable(Action<CKTable> test)
        {
            for (int i = 1; i <= Math.Min(5, _testFile.Tables.Count); i++)
            {
                var table = _testFile.Tables[i];
                TestContext.WriteLine($"Running on table index {i}");
                test(table);
            }
        }

        [TestMethod]
        public void Column_Indexer_AccessesValidCell()
        {
            RunOnEachTestTable(table =>
            {
                var column = table.Columns[1];
                var cell = column[1];
                Assert.IsNotNull(cell);
                Assert.AreEqual(1, cell.CellRef.ColumnIndex);
            });
        }

        [TestMethod]
        public void Column_Indexer_ThrowsOutOfRange()
        {
            RunOnEachTestTable(table =>
            {
                var column = table.Columns[1];
                Assert.ThrowsException<ArgumentOutOfRangeException>(() =>
                {
                    var cell = column[0];
                });
            });
        }

        [TestMethod]
        public void Column_CorrectlyInitializesCellRef()
        {
            RunOnEachTestTable(table =>
            {
                var column = table.Columns[1];
                var refData = column.CellRef;
                Assert.AreEqual(1, refData.Index);
                Assert.AreEqual(table, refData.Table);
            });
        }

        [TestMethod]
        public void CKColumns_Indexer_ReturnsExpectedColumn()
        {
            RunOnEachTestTable(table =>
            {
                var columns = table.Columns;
                var column = columns[1];
                Assert.IsNotNull(column);
                Assert.AreEqual(1, column.Index);
            });
        }

        [TestMethod]
        public void CKColumns_Count_MatchesCOMColumnCount()
        {
            RunOnEachTestTable(table =>
            {
                int expected = table.COMTable.Columns.Count;
                int actual = table.Columns.Count;
                Assert.AreEqual(expected, actual);
            });
        }

        [TestMethod]
        public void Column_SlowDelete_ExecutesSafely()
        {
            RunOnEachTestTable(table =>
            {
                var column = table.Columns[1];
                try
                {
                    column.SlowDelete();
                    Assert.IsTrue(true, "SlowDelete executed without exception.");
                }
                catch (Exception ex)
                {
                    Assert.Fail($"SlowDelete threw exception: {ex.Message}");
                }
            });


        }
        [TestMethod]
        public void TableAccessMode_IncludeAllCells_ReturnsAll()
        {
            var table = _testFile.Tables[3];
            var originalMode = table.AccessMode;
            var testCol = 2;

            try
            {
                table.AccessMode = TableAccessMode.IncludeAllCells;
                var count = table.Columns[testCol].Count;
                TestContext.WriteLine($"IncludeAllCells count: {count}");

                Assert.IsTrue(count > 0, "Expected at least one cell when including all.");
            }
            finally
            {
                table.AccessMode = originalMode;
            }
        }
        [TestMethod]
        public void TableAccessMode_IncludeOnlyAnchorCells_ExcludesMergedFollowers()
        {
            var table = _testFile.Tables[3];
            var originalMode = table.AccessMode;
            var testCol = 2;

            try
            {
                table.AccessMode = TableAccessMode.IncludeAllCells;
                var totalCount = table.Columns[testCol].Count;

                table.AccessMode = TableAccessMode.IncludeOnlyAnchorCells;
                var anchorCount = table.Columns[testCol].Count;

                TestContext.WriteLine($"Total: {totalCount}, Anchors Only: {anchorCount}");

                Assert.IsTrue(anchorCount <= totalCount, "Expected fewer or equal anchor cells.");
                Assert.IsTrue(anchorCount < totalCount, "Expected merged followers to be excluded.");
            }
            finally
            {
                table.AccessMode = originalMode;
            }
        }
        [TestMethod]
        public void TableAccessMode_ExcludeAllMergedCells_ExcludesMergedAndMasters()
        {
            var table = _testFile.Tables[3];
            var originalMode = table.AccessMode;
            var testCol = 2;

            try
            {
                table.AccessMode = TableAccessMode.IncludeAllCells;
                var totalCount = table.Columns[testCol].Count;

                table.AccessMode = TableAccessMode.ExcludeAllMergedCells;
                var dataOnlyCount = table.Columns[testCol].Count;

                TestContext.WriteLine($"Total: {totalCount}, Data-Only: {dataOnlyCount}");

                Assert.IsTrue(dataOnlyCount <= totalCount, "Expected fewer or equal non-merged cells.");
                Assert.IsTrue(dataOnlyCount < totalCount, "Expected all merged cells (including masters) to be excluded.");
            }
            finally
            {
                table.AccessMode = originalMode;
            }
        }

    }
}
