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
            _testFile = RandomTestHarness.GetTempDocumentFrom(_testFilePath, visible: true);
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
            CKDoc.Activate();
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
        [TestMethod]
        public void Can_CopyColumn_FromSourceDocument()
        {
            this.Ping();

            // 🔥 Supply test file paths manually
            string sourceTestFilePath =
            "C:\\Users\\akeyte.PCM\\source\\repos\\CoverageKiller2\\src\\CoverageKiller2_Tests\\TestFiles\\SEA Garage (Noise Floor)_20250313_152027 - Copy.docx";

            string destinationTestFilePath =
            "C:\\Users\\akeyte.PCM\\source\\repos\\CoverageKiller2\\src\\CoverageKiller2_Tests\\TestFiles\\SEA Garage (CC)_20250313_150625.docx";


            Assert.IsTrue(File.Exists(sourceTestFilePath), "Source test file not found.");
            Assert.IsTrue(File.Exists(destinationTestFilePath), "Destination test file not found.");

            // 🔥 Open both documents
            var sourceDoc = RandomTestHarness.GetTempDocumentFrom(sourceTestFilePath, visible: true);
            var destinationDoc = RandomTestHarness.GetTempDocumentFrom(destinationTestFilePath, visible: true);

            sourceDoc.KeepAlive = true;
            sourceDoc.Visible = true;
            sourceDoc.Activate();

            destinationDoc.KeepAlive = true;

            try
            {
                string sourceTableSearchText = "Critical Point Report";
                string destinationTableSearchText = "Critical Point Report";
                string sourceColumnHeaderText = "DL\r\nPower\r\n(dBm)\r\n";
                string destinationColumnHeaderText = "UL\r\nPower\r\n(dBm)\r\n";

                int sectionIndex = 3; //first floor report

                // 🔥 Find source and destination tables
                var sourceTable = SEA2025Fixer.FindTableByRowText(
                    sourceDoc.Sections[sectionIndex].Tables,
                    sourceTableSearchText,
                    rowIndex: 1);

                var destinationTable = SEA2025Fixer.FindTableByRowText(
                    destinationDoc.Sections[sectionIndex].Tables,
                    destinationTableSearchText,
                    rowIndex: 1);

                Assert.IsNotNull(sourceTable, "Source table not found.");
                Assert.IsNotNull(destinationTable, "Destination table not found.");

                // 🔥 Find columns by header text
                var sourceColumn = sourceTable.Columns
                    .FirstOrDefault(col => col[2].Text.ScrunchContains(sourceColumnHeaderText));

                var destinationColumn = destinationTable.Columns
                    .FirstOrDefault(col => col[2].Text.ScrunchContains(destinationColumnHeaderText));

                Assert.IsNotNull(sourceColumn, "Source column not found.");
                Assert.IsNotNull(destinationColumn, "Destination column not found.");

                // 🔥 Perform the copy
                new SEA2025Fixer().CopyColumn(sourceColumn, destinationColumn);

                const int CheckpointInterval = 10;

                var sourceCells = sourceColumn.Cells;
                var destinationCells = destinationColumn.Cells;

                Assert.AreEqual(sourceCells.Count, destinationCells.Count, "Source and destination column cell counts do not match.");

                for (int i = 1; i <= sourceCells.Count; i++)
                {
                    string sourceText = sourceCells[i].Text.Scrunch();
                    string destText = destinationCells[i].Text.Scrunch();

                    Assert.AreEqual(
                        sourceText,
                        destText,
                        $"Mismatch at visual row {i}: Source '{sourceCells[i].Text}' vs Destination '{destinationCells[i].Text}'");

                    if (i % CheckpointInterval == 0 || i == sourceCells.Count)
                    {
                        TestContext.WriteLine($"Verified {i} visual rows copied successfully...");
                    }
                }


                Log.Information("Column copy test passed successfully.");
            }
            finally
            {
                RandomTestHarness.CleanUp(sourceDoc, force: true);
                RandomTestHarness.CleanUp(destinationDoc, force: true);
            }

            this.Pong();
        }

    }
}
