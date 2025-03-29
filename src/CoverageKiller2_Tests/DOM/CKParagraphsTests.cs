using CoverageKiller2.Tests;  // Contains LiveWordDocument helper.
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Diagnostics;

namespace CoverageKiller2.DOM

{
    [TestClass]
    public class CKParagraphsTests
    {
        [TestMethod]
        public void CKParagraphs_Count_MatchesUnderlyingWordParagraphsCount()
        {
            LiveWordDocument.WithTestDocument(LiveWordDocument.DefaultTestFile, doc =>
            {
                // Create a CKRange from the document content.
                CKRange range = new CKRange(doc.Content);
                CKParagraphs paragraphs = range.Paragraphs;

                // Expected count is taken from the underlying Word document's Paragraphs.
                int expectedCount = doc.Paragraphs.Count;
                Assert.AreEqual(expectedCount, paragraphs.Count,
                    "CKParagraphs count should match the Word document's Paragraphs count.");
            });
        }

        [TestMethod]
        public void CKParagraphs_Indexer_ReturnsValidCKParagraph_And_ThrowsOnInvalidIndex()
        {
            LiveWordDocument.WithTestDocument(LiveWordDocument.DefaultTestFile, doc =>
            {
                CKRange range = new CKRange(doc.Content);
                CKParagraphs paragraphs = range.Paragraphs;

                // If there is at least one paragraph, verify that a valid index returns a CKParagraph.
                if (paragraphs.Count > 0)
                {
                    CKParagraph para1 = paragraphs[1];
                    Assert.IsNotNull(para1, "CKParagraph returned by a valid index should not be null.");
                }

                // Index 0 should throw an ArgumentOutOfRangeException.
                Assert.ThrowsException<ArgumentOutOfRangeException>(() =>
                {
                    var p = paragraphs[0];
                }, "Accessing index 0 should throw an ArgumentOutOfRangeException.");

                // Accessing an index greater than Count should throw an exception.
                Assert.ThrowsException<ArgumentOutOfRangeException>(() =>
                {
                    var p = paragraphs[paragraphs.Count + 1];
                }, "Accessing an index greater than Count should throw an ArgumentOutOfRangeException.");
            });
        }

        //[TestMethod]
        public void CKParagraphs_Enumeration_YieldsAllParagraphs()
        {
            LiveWordDocument.WithTestDocument(LiveWordDocument.DefaultTestFile, doc =>
            {
                CKRange range = new CKRange(doc.Content);
                CKParagraphs paragraphs = range.Paragraphs;

                // Cache the count at the start of enumeration.
                int fixedCount = paragraphs.Count;
                int enumeratedCount = 0;
                foreach (CKParagraph para in paragraphs)
                {
                    enumeratedCount++;
                }

                Assert.AreEqual(fixedCount, enumeratedCount,
                    "Enumeration of CKParagraphs should yield the same number of items as the Count property at the start.");
            });
        }
        //[TestMethod]
        public void CKParagraphs_Enumeration_PerformanceMetrics()
        {
            int batchSize = 100; // Adjust this value as needed.
            LiveWordDocument.WithTestDocument(LiveWordDocument.DefaultTestFile, doc =>
            {
                Debug.WriteLine($"Starting Enumeration Performance Test");


                CKRange range = new CKRange(doc.Content);
                CKParagraphs paragraphs = range.Paragraphs;

                int enumeratedCount = 0;
                Stopwatch stopwatch = Stopwatch.StartNew();
                TimeSpan lastBatchTime = stopwatch.Elapsed;

                foreach (CKParagraph para in paragraphs)
                {
                    enumeratedCount++;

                    if (enumeratedCount % batchSize == 0)
                    {
                        TimeSpan currentTime = stopwatch.Elapsed;
                        TimeSpan delta = currentTime - lastBatchTime;
                        Debug.WriteLine($"Enumerated {enumeratedCount} paragraphs; last {batchSize} took {delta.TotalMilliseconds} ms.");
                        lastBatchTime = currentTime;
                    }
                }

                Debug.WriteLine($"Total enumerated paragraphs: {enumeratedCount}");
            });
        }
        [TestMethod]
        public void CKParagraphs_ToString_ReturnsValidString()
        {
            LiveWordDocument.WithTestDocument(LiveWordDocument.DefaultTestFile, doc =>
            {
                CKRange range = new CKRange(doc.Content);
                CKParagraphs paragraphs = range.Paragraphs;
                string output = paragraphs.ToString();

                // Verify that the output contains the expected label and count.
                Assert.IsTrue(output.Contains("CKParagraphs"), "ToString() should contain 'CKParagraphs'.");
                Assert.IsTrue(output.Contains("Count:"), "ToString() should contain 'Count:'.");
                Assert.IsTrue(output.Contains(paragraphs.Count.ToString()), "ToString() should include the count value.");
            });
        }
    }
}
