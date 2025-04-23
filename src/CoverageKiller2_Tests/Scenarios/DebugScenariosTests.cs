using CoverageKiller2.DOM;
using CoverageKiller2.DOM.Tables;
using CoverageKiller2.Pipeline.Processes;
using CoverageKiller2.Test;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Serilog;
using System.Linq;

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
    }
}
