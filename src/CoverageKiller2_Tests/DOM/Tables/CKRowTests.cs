
using CoverageKiller2.Test;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Serilog;
using System;
using System.Diagnostics;

namespace CoverageKiller2.DOM.Tables
{
    [TestClass()]
    public class CKRowTests
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

        private void RunOnEachTestTable(Action<CKTable, TableAccessMode> test)
        {
            var stopwatch = new Stopwatch();
            stopwatch.Start();
            int runCount = 0;

            foreach (var mode in new[] { TableAccessMode.IncludeAllCells, TableAccessMode.IncludeOnlyAnchorCells })
            {
                for (int i = 1; i <= Math.Min(5, _testFile.Tables.Count); i++)
                {
                    runCount++;
                    var table = _testFile.Tables[i];
                    Log.Information($"[START] Row test on table {i}, mode: {mode}");
                    TestContext.WriteLine($"Running on table index {i} with mode: {mode}");

                    var perTest = Stopwatch.StartNew();

                    try
                    {
                        test(table, mode);
                        perTest.Stop();
                        Log.Information($"[PASS] Table {i}, mode: {mode} ({perTest.ElapsedMilliseconds} ms)");
                    }
                    catch (Exception ex)
                    {
                        perTest.Stop();
                        Log.Error(ex, $"[FAIL] Table {i}, mode: {mode} ({perTest.ElapsedMilliseconds} ms)");
                        throw;
                    }

                    var avg = stopwatch.Elapsed.TotalMilliseconds / runCount;
                    var estRemaining = (2 * Math.Min(5, _testFile.Tables.Count) - runCount) * avg;
                    Log.Information($"[ESTIMATE] Average: {avg:F1} ms/test — Estimated time remaining: {estRemaining:F1} ms");
                }
            }

            stopwatch.Stop();
            Log.Information($"[COMPLETE] All tests finished in {stopwatch.Elapsed.TotalSeconds:F2} seconds");
        }

        [TestMethod]
        public void Row_Indexer_AccessesValidCell()
        {
            RunOnEachTestTable((table, mode) =>
            {
                var rowRef = new CKRowCellRef(1, table, table.Rows, mode);
                var row = new CKRow(rowRef, table.Rows);
                var cell = row[1];
                Assert.IsNotNull(cell);
                Assert.AreEqual(1, cell.CellRef.RowIndex);
            });
        }

        [TestMethod]
        public void Row_Indexer_ThrowsOutOfRange()
        {
            RunOnEachTestTable((table, mode) =>
            {
                var rowRef = new CKRowCellRef(1, table, table.Rows, mode);
                var row = new CKRow(rowRef, table.Rows);
                Assert.ThrowsException<ArgumentOutOfRangeException>(() =>
                {
                    var cell = row[0];
                });
            });
        }

        [TestMethod]
        public void Row_CorrectlyInitializesCellRef()
        {
            RunOnEachTestTable((table, mode) =>
            {
                var rowRef = new CKRowCellRef(1, table, table.Rows, mode);
                var row = new CKRow(rowRef, table.Rows);
                Assert.AreEqual(1, rowRef.Index);
                Assert.AreEqual(table, rowRef.Table);
            });
        }

        [TestMethod]
        public void CKRows_Indexer_ReturnsExpectedRow()
        {
            RunOnEachTestTable((table, mode) =>
            {
                var row = table.Rows[1];
                Assert.IsNotNull(row);
                Assert.AreEqual(1, row.RowRef.RowIndex);
            });
        }

        [TestMethod]
        public void CKRows_Count_MatchesCOMRowCount()
        {
            RunOnEachTestTable((table, mode) =>
            {
                int expected = table.COMTable.Rows.Count;
                int actual = table.Rows.Count;
                Assert.AreEqual(expected, actual);
            });
        }

        [TestMethod]
        public void Row_SlowDelete_ExecutesSafely()
        {
            RunOnEachTestTable((table, mode) =>
            {
                var rowRef = new CKRowCellRef(1, table, table.Rows, mode);
                var row = new CKRow(rowRef, table.Rows);

                try
                {
                    row.SlowDelete();
                    Assert.IsTrue(true, "SlowDelete executed without exception.");
                }
                catch (Exception ex)
                {
                    Assert.Fail($"SlowDelete threw exception: {ex.Message}");
                }
            });
        }
    }
}
