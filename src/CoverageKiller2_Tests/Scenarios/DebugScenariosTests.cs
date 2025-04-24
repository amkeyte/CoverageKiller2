using CoverageKiller2.DOM;
using CoverageKiller2.DOM.Tables;
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

            Assert.Fail("Could not find table labeled 'Test Details'");
        }
        [TestMethod]
        public void Edits_AffectOriginalDocument()
        {
            // Step 1: Open the document using the harness
            var doc = RandomTestHarness.GetTempDocumentFrom(RandomTestHarness.TestFile1);
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

    }
}
