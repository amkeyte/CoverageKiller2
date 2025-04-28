using CoverageKiller2.Test;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Serilog;
using System;

namespace CoverageKiller2.DOM
{
    /// <summary>
    /// Unit tests for <see cref="CKParagraphs"/> collection wrapper.
    /// </summary>
    /// <remarks>
    /// Version: CK2.00.02.0000
    /// </remarks>
    [TestClass]
    public class CKParagraphsTests
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
        public void CKParagraphs_Count_MatchesUnderlyingWordParagraphsCount()
        {
            var range = _testFile.Sections[1];
            var paragraphs = range.Paragraphs;

            int expectedCount = range.COMRange.Paragraphs.Count;
            Assert.AreEqual(expectedCount, paragraphs.Count,
                "CKParagraphs count should match the Word Section's Paragraphs count.");
        }

        [TestMethod]
        public void CKParagraphs_Indexer_ReturnsValidCKParagraph_And_ThrowsOnInvalidIndex()
        {
            var range = _testFile.Sections[1];
            var paragraphs = range.Paragraphs;

            if (paragraphs.Count > 0)
            {
                var para1 = paragraphs[1];
                Assert.IsNotNull(para1, "CKParagraph returned by a valid index should not be null.");
            }

            Assert.ThrowsException<ArgumentOutOfRangeException>(() =>
            {
                var p = paragraphs[0];
            }, "Accessing index 0 should throw an ArgumentOutOfRangeException.");

            Assert.ThrowsException<ArgumentOutOfRangeException>(() =>
            {
                var p = paragraphs[paragraphs.Count + 1];
            }, "Accessing an index greater than Count should throw an ArgumentOutOfRangeException.");
        }

        [TestMethod]
        public void CKParagraphs_Enumeration_YieldsAllParagraphs()
        {
            var range = _testFile.Sections[1];
            var paragraphs = range.Paragraphs;

            int fixedCount = paragraphs.Count;
            int enumeratedCount = 0;

            Log.Verbose($"Enumerating {fixedCount} paragraphs.");
            foreach (var para in paragraphs)
            {
                enumeratedCount++;
                if (enumeratedCount % 20 == 0) Log.Verbose($"_____ {enumeratedCount} paragraphs iterated _____");
            }

            Assert.AreEqual(fixedCount, enumeratedCount,
                "Enumeration of CKParagraphs should yield the same number of items as the Count property at the start.");
        }

        [TestMethod]
        public void CKParagraphs_ToString_ReturnsValidString()
        {
            var range = _testFile.Sections[1];
            var paragraphs = range.Paragraphs;
            string output = paragraphs.ToString();

            Assert.IsTrue(output.Contains("CKParagraphs"), "ToString() should contain 'CKParagraphs'.");
            Assert.IsTrue(output.Contains("Count:"), "ToString() should contain 'Count:'.");
            Assert.IsTrue(output.Contains(paragraphs.Count.ToString()), "ToString() should include the count value.");
        }

        #region DeferCOM Tests

        [TestMethod]
        public void CKParagraphs_DeferConstructor_StartsDirty()
        {
            var range = _testFile.Sections[1];
            var comParas = range.COMRange.Paragraphs;
            var deferredParas = new CKParagraphs(comParas, range, deferCOM: true);

            Assert.IsTrue(deferredParas.IsDirty, "Deferred CKParagraphs should initially be dirty.");
        }

        [TestMethod]
        public void CKParagraphs_DeferConstructor_EnumerationLiftsDefer()
        {
            var range = _testFile.Sections[1];
            var comParas = range.COMRange.Paragraphs;
            var deferredParas = new CKParagraphs(comParas, range, deferCOM: true);

            Assert.IsTrue(deferredParas.IsDirty, "Deferred CKParagraphs should start dirty.");

            int enumeratedCount = 0;
            foreach (var para in deferredParas)
            {
                // touching paragraphs should lift defer automatically through Cache
                enumeratedCount++;
            }

            Assert.IsFalse(deferredParas.IsDirty, "After enumeration, CKParagraphs should no longer be dirty.");
            Assert.AreEqual(comParas.Count, enumeratedCount, "Enumeration should yield all paragraphs after lifting defer.");
        }

        [TestMethod]
        public void CKParagraphs_ManualRefresh_RebuildsParagraphList()
        {
            var range = _testFile.Sections[1];
            var comParas = range.COMRange.Paragraphs;
            var deferredParas = new CKParagraphs(comParas, range, deferCOM: true);

            Assert.IsTrue(deferredParas.IsDirty, "Deferred CKParagraphs should start dirty.");

            deferredParas.Refresh();

            Assert.IsFalse(deferredParas.IsDirty, "After Refresh(), CKParagraphs should no longer be dirty.");
            Assert.AreEqual(comParas.Count, deferredParas.Count, "Paragraph count after Refresh() should match Word COM count.");
        }

        #endregion
    }
}
