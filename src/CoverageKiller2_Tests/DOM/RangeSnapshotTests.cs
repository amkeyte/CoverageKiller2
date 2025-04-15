using CoverageKiller2.DOM;
using CoverageKiller2.Test;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;

namespace CoverageKiller2.DOM
{
    /// <summary>
    /// Tests for verifying that <see cref="RangeSnapshot"/> detects changes in Word ranges.
    /// </summary>
    /// <remarks>
    /// Version: CK2.00.00.0000
    /// </remarks>
    [TestClass]
    public class RangeSnapshotTests
    {
        private CKDocument _doc;

        [TestInitialize]
        public void SetUp()
        {
            _doc = RandomTestHarness.GetTempDocumentFrom(RandomTestHarness.TestFile1);
        }

        [TestCleanup]
        public void TearDown()
        {
            RandomTestHarness.CleanUp(_doc);
            _doc = null;
        }

        [TestMethod]
        public void RangeSnapshots_ShouldDetectChangesAfterInsertion()
        {
            var range = _doc.Range();
            var paragraphs = range.Paragraphs;
            var tables = range.Tables;

            var ranges = new List<CKRange>
            {

                paragraphs.Count >= 1 ? paragraphs[1] : null,
                paragraphs.Count >= 2 ? paragraphs[2] : null,
                tables.Count > 0 ? tables[1] : null,
                _doc.Range(0, 10)
            }.Where(r => r != null).ToList();

            Assert.IsTrue(ranges.Count > 0, "No usable ranges found to snapshot.");

            var snapshotsBefore = ranges.Select(r => new RangeSnapshot(r.COMRange)).ToList();

            _doc.Range(0, 0).COMRange.InsertBefore("PREPENDED TEXT. ");

            var snapshotsAfter = ranges.Select(r => new RangeSnapshot(r.COMRange)).ToList();

            bool anyChanged = false;
            for (int i = 0; i < snapshotsBefore.Count; i++)
            {
                var before = snapshotsBefore[i];
                var after = snapshotsAfter[i];
                bool match = before.FastMatch(after);

                Debug.WriteLine("");
                Debug.WriteLine($"Range {i}: {(match ? "UNCHANGED" : "CHANGED")}");
                Debug.WriteLine($"Before: {before.FastHash} - After: {after.FastHash}");
                Debug.WriteLine($"TextBefore: {before.TextPreview}");
                Debug.WriteLine($"TextAfter: {after.TextPreview}");

                if (!match) anyChanged = true;
            }

            Assert.IsTrue(anyChanged, "Expected at least one snapshot to detect a change.");
        }
    }
}
