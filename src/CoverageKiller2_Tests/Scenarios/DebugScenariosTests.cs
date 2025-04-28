using CoverageKiller2.DOM;
using CoverageKiller2.DOM.Tables;
using CoverageKiller2.Logging;
using CoverageKiller2.Pipeline.Processes;
using CoverageKiller2.Test;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Serilog;
using System;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;

namespace CoverageKiller2.Tests.Scenarios

{
    [TestClass]
    public class DebugScenariosTests
    {
        private string _testFilePath;
        private CKDocument _testFile;

        public TestContext TestContext { get; set; }

        [TestInitialize]
        public void Setup()
        {
            Log.Information($"Running test => {GetType().Name}::{TestContext.TestName}");
            _testFilePath = RandomTestHarness.TestFile2;
            _testFile = RandomTestHarness.GetTempDocumentFrom(_testFilePath);
        }

        [TestCleanup]
        public void Cleanup()
        {
            RandomTestHarness.CleanUp(_testFile, force: true);
            Log.Information($"Completed test => {GetType().Name}::{TestContext.TestName}; status: {TestContext.CurrentTestOutcome}");
        }

        [TestMethod]
        public void CanFindTestDetailsTable_ByScrunchedRow1_A()
        {
            var CKDoc = _testFile;
            //CKDoc.KeepAlive = true;
            int rowIndex = 1;
            string TDTable_ss = "Test Details";

            var table = SEA2025Fixer.FindTableByRowText(
                CKDoc.Tables,
                TDTable_ss,
                rowIndex,
                TableAccessMode.IncludeOnlyAnchorCells);

            Assert.IsNotNull(table, "Could not find table labeled 'Test Details'");

            string headerText = string.Join(string.Empty, table.Rows[1].Select(c => c.Text));
            Assert.IsTrue(CKTextHelper.ScrunchEquals(headerText, TDTable_ss), "Scrunched header did not match.");

        }
        [TestMethod]
        public void CanFindTestDetailsTable_ByScrunchedRow1_B()
        {
            var CKDoc = _testFile;
            CKDoc.KeepAlive = true;
            string TDTable_ss = "Test Details";
            string scrunchedTarget = CKTextHelper.Scrunch(TDTable_ss);

            foreach (var table in CKDoc.Tables)
            {
                table.AccessMode = TableAccessMode.IncludeOnlyAnchorCells;
                var row = table.Rows[1]; //fail is here

                var cellTexts = row.Select(c => c.Text).ToList(); // list of text from each cell
                var joinedText = string.Join(string.Empty, cellTexts);
                Log.Debug($"Table Text returned {joinedText}");
                var scrunchedRowText = CKTextHelper.Scrunch(joinedText);

                TestContext.WriteLine($"Table {table.DocumentTableIndex}: Row 1 scrunched text = '{scrunchedRowText}'");

                if (scrunchedRowText == scrunchedTarget)
                {
                    TestContext.WriteLine("Match found.");
                    Assert.IsTrue(true);
                    return;
                }
            }

            Assert.Fail("Could not find table labeled 'Test Details'");
        }
        [TestMethod]
        public void Edits_AffectOriginalDocument()
        {
            var doc = _testFile;
            var comDoc = doc.GiveMeCOMDocumentIWillOwnItAndPromiseToCleanUpAfterMyself();
            doc.KeepAlive = true;
            try
            {
                // Step 2: Validate that the CKDocument wraps the same Word.Document
                Assert.IsTrue(doc.Matches(comDoc), "CKDocument does not match original Word.Document");

                // Step 3: Confirm there's at least one table
                Assert.IsTrue(comDoc.Tables.Count > 0, "No tables in COM Document");
                Assert.IsTrue(doc.Tables.Count > 0, "No tables in CKDocument");

                var tableIndex = 1;
                var table = doc.Tables[tableIndex];
                var colCountBefore = table.COMTable.Columns.Count;

                // Step 4: Perform an edit through CKTable
                var colToDelete = table.Columns[1];
                colToDelete.Delete();

                var colCountAfter = table.COMTable.Columns.Count;

                // Step 5: Confirm column count decreased
                Assert.AreEqual(colCountBefore - 1, colCountAfter, "Column deletion did not take effect on COMTable");

                // Step 6: Confirm same table in both layers
                Assert.IsTrue(doc.Matches(table.COMTable.Range.Document), "Table edit applied to wrong document");

                // Step 7 (optional): Save and re-open to check persistence
                var savedPath = Path.Combine(Path.GetTempPath(), "EditConfirm_" + Guid.NewGuid() + ".docx");
                comDoc.SaveAs2(savedPath);
                var confirmDoc = RandomTestHarness.GetTempDocumentFrom(savedPath);
                var confirmTable = confirmDoc.Tables[tableIndex];

                Assert.AreEqual(colCountAfter, confirmTable.COMTable.Columns.Count, "Saved document does not reflect the column change");
            }
            finally
            {
                // Cleanup
                Marshal.ReleaseComObject(comDoc);
            }
        }
        [TestMethod]
        public void CanFindTestDetailsTable_ByScrunchedRow1_C()
        {
            var CKDoc = _testFile;
            CKDoc.KeepAlive = true;
            string TDTable_ss = "Test Details";
            string scrunchedTarget = CKTextHelper.Scrunch(TDTable_ss);

            foreach (var table in CKDoc.Tables)
            {
                table.AccessMode = TableAccessMode.IncludeOnlyAnchorCells;

                if (table.Rows.Count < 1)
                {
                    TestContext.WriteLine($"Table {table.DocumentTableIndex}: has no rows.");
                    continue;
                }

                var row = table.Rows[1];

                var cellTexts = row.Select(c => c.Text).ToList(); // list of text from each cell
                var joinedText = string.Join(string.Empty, cellTexts);
                Log.Debug($"Table Text returned {joinedText}");
                var scrunchedRowText = CKTextHelper.Scrunch(joinedText);

                TestContext.WriteLine($"Table {table.DocumentTableIndex}: Row 1 scrunched text = '{scrunchedRowText}'");

                if (scrunchedRowText == scrunchedTarget)
                {
                    TestContext.WriteLine("Match found.");
                    Assert.IsTrue(true);
                    return;
                }
            }
        }



        /// <summary>
        /// Bug 20250425-0013
        /// </summary>
        [TestMethod]


        public void CrawlHoriz_HandlesMisalignedGrid_Table4()
        {
            var CKDoc = _testFile;
            CKDoc.KeepAlive = true;

            int tableIndex = 4;
            Assert.IsTrue(CKDoc.Tables.Count >= tableIndex, "Test file does not contain table 4.");

            var table = CKDoc.Tables[tableIndex];
            table.AccessMode = TableAccessMode.IncludeOnlyAnchorCells;

            try
            {
                var rows = table.Rows; // should trigger Grid + CrawlHoriz
                Assert.IsTrue(rows.Count > 0, "Table returned no rows.");
                TestContext.WriteLine("Table 4 was processed successfully.");
            }
            catch (Exception ex)
            {
                Assert.Fail($"CrawlHoriz failed unexpectedly: {ex.Message}");
            }
        }

        /// <summary>
        /// Bug: Issue 1
        /// </summary>
        [TestMethod]
        public void Can_RemoveCriticalPointFields_FromFloorSectionTable()
        {
            Log.Debug("[Issue1] Testing removal of fields from Critical Points Report");
            var CKDoc = _testFile;
            CKDoc.KeepAlive = true;

            // Setup
            string searchText = "Critical Point Report";
            int rowIndex = 1;

            var floorSectionCriticalPointsTable = SEA2025Fixer.FindTableByRowText(
                CKDoc.Tables,
                searchText,
                rowIndex,
                TableAccessMode.IncludeOnlyAnchorCells); // avoid repeating header merged cells

            Assert.IsNotNull(floorSectionCriticalPointsTable, "[Issue1]Critical Point Report table not found.");

            var originalColCount = floorSectionCriticalPointsTable.Columns.Count;
            Log.Debug($"[Issue1] Original column count: {originalColCount}");

            // Deletion logic matching your pipeline
            var headersToRemove = "UL\r\nPower\r\n(dBm)\tUL\r\nS/N\r\n(dB)\tUL\r\nFBER\r\n(%)\tResult\tDL\r\nLoss\r\n(dB)\r\n"
                .Split(new[] { '\t' }, StringSplitOptions.RemoveEmptyEntries)
                .Select(s => s.Scrunch())
                .ToList();

            Log.Debug($"[Issue1] Headers to Remove: {headersToRemove.DumpString()}");

            var headersFound = floorSectionCriticalPointsTable.Columns.Select(col => col[2].Text);
            Log.Debug($"[Issue1] Headers found: {headersToRemove.DumpString()}");

            floorSectionCriticalPointsTable.Columns
                .Delete(col => headersToRemove.Contains(col[2].ScrunchedText));

            floorSectionCriticalPointsTable.MakeFullPage(); // finalize layout

            // Recheck
            var updatedColCount = floorSectionCriticalPointsTable.COMTable.Columns.Count;
            Log.Debug($"[Issue1] Updated column count: {updatedColCount}");

            // Assert: At least one column must have been deleted
            Assert.IsTrue(updatedColCount < originalColCount, "[Issue1] No columns were deleted from the Critical Point Report table.");

            // Extra safety: Make sure columns that should have been removed are gone
            foreach (var col in floorSectionCriticalPointsTable.Columns)
            {
                var text = col[2].Text.Scrunch();
                Assert.IsFalse(headersToRemove.Contains(text), $"[Issue1] Unexpected leftover column header: '{text}'");
            }
        }

    }
}
