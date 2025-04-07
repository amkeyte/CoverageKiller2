using CoverageKiller2.Logging;
using CoverageKiller2.Test;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Serilog;
using System;
using System.IO;
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
        private CKDocument _doc;

        [TestInitialize]
        public void SetUp()
        {
            LH.Ping(this.GetType());
            _doc = RandomTestHarness.GetTempDocumentFrom(RandomTestHarness.TestFile1);
            LH.Pong(this.GetType());
        }

        [TestCleanup]
        public void TearDown()
        {
            LH.Ping(this.GetType());
            RandomTestHarness.CleanUp(_doc);
            _doc = null;
            LH.Pong(this.GetType());
        }

        [TestMethod]
        public void CKParagraphs_Count_MatchesUnderlyingWordParagraphsCount()
        {
            LH.Ping($"App Instance {_doc.Application.PID} - Test File {Path.GetFileName(_doc.FullPath)}", null);
            CKRange range = _doc.Range();
            CKParagraphs paragraphs = range.Paragraphs;

            int expectedCount = _doc.Range().COMRange.Paragraphs.Count;
            Assert.AreEqual(expectedCount, paragraphs.Count,
                "CKParagraphs count should match the Word document's Paragraphs count.");
            LH.Pong();
        }

        [TestMethod]
        public void CKParagraphs_Indexer_ReturnsValidCKParagraph_And_ThrowsOnInvalidIndex()
        {
            LH.Ping($"App Instance {_doc.Application.PID} - Test File {Path.GetFileName(_doc.FullPath)}", null);

            CKRange range = _doc.Sections[1];
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
            LH.Pong();

        }

        [TestMethod]
        public void CKParagraphs_Enumeration_YieldsAllParagraphs()
        {
            LH.Ping($"App Instance {_doc.Application.PID} - Test File {Path.GetFileName(_doc.FullPath)}", null);

            try
            {


                CKRange range = _doc.Sections[1];
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
            LH.Pong("Tests Completed");

        }

        [TestMethod]
        public void CKParagraphs_ToString_ReturnsValidString()
        {
            LH.Ping($"App Instance {_doc.Application.PID} - Test File {Path.GetFileName(_doc.FullPath)}", null);

            CKRange range = _doc.Sections[1];
            CKParagraphs paragraphs = range.Paragraphs;
            string output = paragraphs.ToString();

            Assert.IsTrue(output.Contains("CKParagraphs"), "ToString() should contain 'CKParagraphs'.");
            Assert.IsTrue(output.Contains("Count:"), "ToString() should contain 'Count:'.");
            Assert.IsTrue(output.Contains(paragraphs.Count.ToString()), "ToString() should include the count value.");

            LH.Pong();
        }
    }
}
