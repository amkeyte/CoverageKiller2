using CoverageKiller2.DOM;
using CoverageKiller2.Test;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Serilog;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;

namespace CoverageKiller2.DOM
{
    /// <summary>
    /// Tests for verifying that <see cref="RangeSnapshot"/> detects changes in Word ranges.
    /// </summary>
    /// <remarks>
    /// Version: CK2.00.00.0001
    /// </remarks>
    [TestClass]
    public class RangeSnapshotTests
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
        public void RangeSnapshots_ShouldDetectChangesAfterInsertion()
        {
            var range = _testFile.Sections[1];
            var paragraphs = range.Paragraphs;
            var tables = range.Tables;

            var labeledRanges = new Dictionary<string, CKRange>
            {
                { "Paragraph[1]", paragraphs.Count >= 1 ? paragraphs[1] : null },
                { "Paragraph[2]", paragraphs.Count >= 2 ? paragraphs[2] : null },
                { "Table[1]", tables.Count > 0 ? tables[1] : null },
                { "DocStart(0-10)", _testFile.Range(0, 10) }
            }.Where(pair => pair.Value != null).ToList();

            Assert.IsTrue(labeledRanges.Count > 0, "No usable ranges found to snapshot.");

            var snapshotsBefore = labeledRanges
                .ToDictionary(pair => pair.Key, pair => new RangeSnapshot(pair.Value.COMRange));

            // Modify the document
            _testFile.Range(0, 0).COMRange.InsertBefore("PREPENDED TEXT. ");

            var snapshotsAfter = labeledRanges
                .ToDictionary(pair => pair.Key, pair => new RangeSnapshot(pair.Value.COMRange));

            bool anyChanged = false;

            foreach (var key in snapshotsBefore.Keys)
            {
                var before = snapshotsBefore[key];
                var after = snapshotsAfter[key];
                bool match = RangeSnapshot.FastMatch(before, after);

                Debug.WriteLine($"--- {key} ---");
                Debug.WriteLine($"Status  : {(match ? "UNCHANGED" : "CHANGED")}");
                Debug.WriteLine($"Hash    : {before.FastHash} => {after.FastHash}");
                Debug.WriteLine($"Preview : \"{before.TextPreview}\" => \"{after.TextPreview}\"");
                Debug.WriteLine("");

                if (!match) anyChanged = true;
            }

            Assert.IsTrue(anyChanged, "Expected at least one snapshot to detect a change.");
        }
    }
}
