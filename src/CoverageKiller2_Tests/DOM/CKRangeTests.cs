using CoverageKiller2.Tests;  // Contains LiveWordDocument helper.
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.Runtime.InteropServices;

namespace CoverageKiller2.DOM
{
    [TestClass]
    public class CKRangeTests
    {


        /// <summary>
        /// Verifies that setting the Text property on a CKRange that spans a partial table
        /// and regular text throws a COMException.
        /// </summary>
        [TestMethod]
        public void CKRange_SetText_OnMixedRange_ThrowsCOMException()
        {
            LiveWordDocument.WithTestDocument(LiveWordDocument.DefaultTestFile, doc =>
            {
                // Create a range that likely spans mixed content (partial table and text).
                var tbl = doc.Range().Tables[1];
                var start = tbl.Range.Start + 3;
                var end = tbl.Range.End + 20;
                CKRange brokenRange = new CKRange(doc.Range(start, end));

                // Attempt to set the Text property, expecting a COMException.
                Assert.ThrowsException<COMException>(() =>
                {
                    brokenRange.COMRange.Text = "Test new text";
                }, "Setting Text on a mixed-content range should throw COMException.");
            });
        }



        /// <summary>
        /// Verifies that when the underlying COMRange text changes, Refresh updates the caches and resets IsDirty.
        /// </summary>
        [TestMethod]
        public void CKRange_Refresh_UpdatesCachesAndResetsDirtyFlag()
        {
            LiveWordDocument.WithTestDocument(LiveWordDocument.DefaultTestFile, doc =>
            {
                // Create a CKRange from the document's content.
                CKRange range = new CKRange(doc.Paragraphs[20].Range);

                // Cache initial values.
                string raw1 = range.Text;
                string pretty1 = range.PrettyText;
                string scrunched1 = range.ScrunchedText;

                // Simulate a change to the underlying COMRange.
                // Note: Modifying COMRange.Text updates the document. In a test environment,
                // ensure that the document is disposable or working on a copy.
                string newRaw = raw1 + " extra";

                range.Text = newRaw;

                // Now the range should be dirty.
                Assert.IsTrue(range.IsDirty, "Range should be dirty after modifying COMRange.Text.");

                // Call Refresh() to update the caches.
                range.Refresh();

                // After refresh, IsDirty should be false.
                Assert.IsFalse(range.IsDirty, "Range should not be dirty after refresh.");

                // Verify that the caches have been updated.
                Assert.IsTrue(CKTextHelper.ScrunchEquals(newRaw, range.Text), "Raw text should be updated.");
                Assert.IsTrue(CKTextHelper.ScrunchEquals(newRaw, range.PrettyText), "Pretty text should be updated.");
                Assert.IsTrue(CKTextHelper.ScrunchEquals(newRaw, range.ScrunchedText), "Scrunched text should be updated.");
            });
        }

        /// <summary>
        /// Verifies that TextEquals compares the scrunched (whitespace-removed) versions of the texts.
        /// </summary>
        [TestMethod]
        public void CKRange_TextEquals_IgnoresWhitespaceDifferences()
        {
            LiveWordDocument.WithTestDocument(LiveWordDocument.DefaultTestFile, doc =>
            {
                CKRange range = new CKRange(doc.Range(20, 100));
                // Get the raw text.
                string raw = range.COMRange.Text;
                // Create a modified string that has extra whitespace.
                string modified = raw + "  \t\n";

                // TextEquals compares scrunched versions, so these should be equal.
                bool areEqual = range.TextEquals(modified);
                Assert.IsTrue(areEqual, "TextEquals should consider texts equal when only whitespace differs.");
            });
        }

        /// <summary>
        /// Verifies that PrettyText properly processes control characters.
        /// Specifically, it should replace cell markers (\a) with tabs, preserve CR+LF sequences,
        /// and remove extraneous control characters.
        /// </summary>
        [TestMethod]
        public void CKRange_PrettyText_ProcessesControlCharactersCorrectly()
        {
            LiveWordDocument.WithTestDocument(LiveWordDocument.DefaultTestFile, doc =>
            {
                // For testing, assign a known sample text to the COMRange.
                // NOTE: This test assumes that modifying COMRange.Text is acceptable for your test document.

                //*** Leaving this test as fail for now in case the \a thing becomes a problem where 
                // since word doesn't have a table in place it just eliminates all control characters.
                // (apparently)

                string sample = "Hello\r\nWorld\aNext";
                CKRange range = new CKRange(doc.Range(20, 100), null);
                // fails right now because of internal Word handling with \a Fix if needed
                range.Text = sample;
                range.Refresh();
                var x = range.Text;
                string pretty = range.PrettyText;
                // Expected result: \a is replaced with a tab, CR+LF is preserved, and extraneous control characters are removed.
                string expected = CKTextHelper.Pretty(sample);
                Assert.AreEqual(expected, pretty, "PrettyText should transform control characters as expected.");
            });
        }
    }
}
