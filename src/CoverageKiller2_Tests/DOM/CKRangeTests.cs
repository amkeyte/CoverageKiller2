using CoverageKiller2.DOM;
using CoverageKiller2.DOM.Tables;
using CoverageKiller2.Test;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.Linq;
using System.Runtime.InteropServices;

namespace CoverageKiller2.Tests.DOM
{
    /// <summary>
    /// Unit tests for the CKRange class.
    /// </summary>
    [TestClass]
    // Version: CK2.00.00.0001
    public class CKRangeTests
    {
        //******* Standard Rigging ********
        private string _testFilePath;
        private CKDocument _testDoc;

        [TestInitialize]
        public void Setup()
        {
            _testFilePath = RandomTestHarness.TestFile1;
            _testDoc = RandomTestHarness.GetTempDocumentFrom(_testFilePath);
        }
        [TestCleanup]
        public void Cleanup()
        {
            RandomTestHarness.CleanUp(_testDoc, force: true);
        }
        //******* End Standard Rigging ********



        /// <summary>
        /// Verifies that setting the Text property on a CKRange that spans a partial table
        /// and regular text throws a COMException.
        /// </summary>
        [TestMethod]
        public void CKRange_SetText_OnMixedRange_ThrowsCOMException()
        {
            CKTable table = _testDoc.Tables.FirstOrDefault();
            Assert.IsNotNull(table, "Test document must contain at least one table.");

            int start = table.Start + 3;
            int end = table.End + 20;

            var mixedRange = _testDoc.Range(start, end);
            //var ckRange = new CKRange(mixedRange.COMRange); // okay to use COMRange if already wrapped

            Assert.ThrowsException<COMException>(() =>
            {
                mixedRange.Text = "Test new text";
            }, "Setting Text on a mixed-content range should throw COMException.");
        }

        /// <summary>
        /// Verifies that when the underlying COMRange text changes, Refresh updates the caches and resets IsDirty.
        /// </summary>
        [TestMethod]
        public void CKRange_Refresh_UpdatesCachesAndResetsDirtyFlag()
        {
            var ckRange = _testDoc.Range(30, 40);

            string original = ckRange.Text;
            string newText = original + " extra";

            ckRange.Text = newText;

            Assert.IsTrue(ckRange.IsDirty, "Range should be dirty after text change.");
            ckRange.Refresh();
            Assert.IsFalse(ckRange.IsDirty, "Range should be clean after refresh.");

            Assert.AreEqual(newText, ckRange.Text);
            //Assert.AreEqual(newText, ckRange.PrettyText);
            Assert.AreEqual(CKTextHelper.Scrunch(newText), ckRange.ScrunchedText);
        }

        /// <summary>
        /// Verifies that TextEquals compares the scrunched (whitespace-removed) versions of the texts.
        /// </summary>
        [TestMethod]
        public void CKRange_TextEquals_IgnoresWhitespaceDifferences()
        {
            var range = _testDoc.Range(20, 100);
            var ckRange = new CKRange(range.COMRange);
            var modified = ckRange.Text + "   \t\n ";

            Assert.IsTrue(ckRange.TextEquals(modified), "Whitespace-only differences should be ignored.");
        }

        ///// <summary>
        ///// Verifies that PrettyText properly processes control characters.
        ///// </summary>
        //[TestMethod]
        //public void CKRange_PrettyText_ProcessesControlCharactersCorrectly()
        //{
        //    string rawText = "Hello\r\nWorld\aNext";

        //    var range = _testDoc.Range(0, 0); // empty range to inject text
        //    var ckRange = new CKRange(range.COMRange);

        //    ckRange.Text = rawText;
        //    ckRange.Refresh();

        //    string expected = CKTextHelper.Pretty(rawText);
        //    Assert.AreEqual(expected, ckRange.PrettyText, "PrettyText did not transform control characters as expected.");
        //}
    }
}
