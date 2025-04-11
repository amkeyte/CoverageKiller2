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
    /// Version: CK2.00.00.0000
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
            CKRange range = _testFile.Range();
            CKParagraphs paragraphs = range.Sections[1].Paragraphs;

            int expectedCount = _testFile.Range().COMRange.Sections[1].Range.Paragraphs.Count;
            Assert.AreEqual(expectedCount, paragraphs.Count,
                "CKParagraphs count should match the Word document's Paragraphs count.");
        }

        [TestMethod]
        public void CKParagraphs_Indexer_ReturnsValidCKParagraph_And_ThrowsOnInvalidIndex()
        {

            CKRange range = _testFile.Sections[1];
            CKParagraphs paragraphs = range.Paragraphs;

            if (paragraphs.Count > 0)
            {
                CKParagraph para1 = paragraphs[1];
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

            try
            {


                CKRange range = _testFile.Sections[1];
                CKParagraphs paragraphs = range.Paragraphs;

                int fixedCount = paragraphs.Count;
                int enumeratedCount = 0;
                Log.Verbose($"Enumerating {fixedCount} paragraphs.");
                foreach (CKParagraph para in paragraphs)
                {
                    enumeratedCount++;
                    if (enumeratedCount % 20 == 0) Log.Verbose($"_____ {enumeratedCount} _____");
                }


                Assert.AreEqual(fixedCount, enumeratedCount,
                    "Enumeration of CKParagraphs should yield the same number of items as the Count property at the start.");
            }
            catch (Exception ex)
            {
                {
                    Log.Error(ex, $"Test Failed. {ex.Message}");
                }

            }

        }

        [TestMethod]
        public void CKParagraphs_ToString_ReturnsValidString()
        {

            CKRange range = _testFile.Sections[1];
            CKParagraphs paragraphs = range.Paragraphs;
            string output = paragraphs.ToString();

            Assert.IsTrue(output.Contains("CKParagraphs"), "ToString() should contain 'CKParagraphs'.");
            Assert.IsTrue(output.Contains("Count:"), "ToString() should contain 'Count:'.");
            Assert.IsTrue(output.Contains(paragraphs.Count.ToString()), "ToString() should include the count value.");

        }
    }
}
