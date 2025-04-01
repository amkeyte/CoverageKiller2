using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using Word = Microsoft.Office.Interop.Word;

namespace CoverageKiller2.DOM
{
    [TestClass]
    public class RangeSnapshotTests
    {
        [TestMethod]
        public void RangeSnapshots_ShouldDetectChangesAfterInsertion()
        {
            LiveWordDocument.WithTestDocument(doc =>
            {
                Word.Document wordDoc = doc.COMDocument;

                // Grab several representative ranges
                var ranges = new List<Word.Range>
                {
                    wordDoc.Paragraphs[1].Range.Duplicate, // First paragraph
                    wordDoc.Paragraphs[2].Range.Duplicate, // Another paragraph
                    wordDoc.Tables.Count > 0 ? wordDoc.Tables[1].Range.Duplicate : null,
                    wordDoc.Range(0, 10).Duplicate          // Range at start of document
                }.Where(r => r != null).ToList();

                // Snapshot BEFORE change
                var snapshotsBefore = ranges.Select(r => new RangeSnapshot(r)).ToList();

                // Insert text at beginning
                wordDoc.Range(0, 0).InsertBefore("PREPENDED TEXT. ");

                // Snapshot AFTER change
                var snapshotsAfter = ranges.Select(r => new RangeSnapshot(r)).ToList();

                // Compare hashes
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
                }

                // This is a sanity check: at least 1 hash should change
                bool anyChanged = snapshotsBefore
                    .Zip(snapshotsAfter, (b, a) => !b.FastMatch(a))
                    .Any(x => x);

                Assert.IsTrue(anyChanged, "Expected at least one snapshot to detect a change.");
            });
        }
    }
}
