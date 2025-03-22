using CoverageKiller2.Tests;  // Contains LiveWordDocument helper.
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;

namespace CoverageKiller2.DOM
{
    [TestClass]
    public class CKSectionsTests
    {
        [TestMethod]
        public void CKSections_Count_MatchesUnderlyingWordSectionsCount()
        {
            LiveWordDocument.WithTestDocument(LiveWordDocument.Default, doc =>
            {
                // Create a CKRange from the document content.
                CKRange range = new CKRange(doc.Content);
                CKSections sections = range.Sections;

                // The expected count is taken from the underlying Word document.
                int expectedCount = doc.Sections.Count;
                Assert.AreEqual(expectedCount, sections.Count,
                    "CKSections count should match the Word document's Sections count.");
            });
        }

        [TestMethod]
        public void CKSections_Indexer_ReturnsValidCKSection_And_ThrowsOnInvalidIndex()
        {
            LiveWordDocument.WithTestDocument(LiveWordDocument.Default, doc =>
            {
                CKRange range = new CKRange(doc.Content);
                CKSections sections = range.Sections;

                // If there is at least one section, check a valid index.
                if (sections.Count > 0)
                {
                    CKSection section1 = sections[1];
                    Assert.IsNotNull(section1, "CKSection returned by a valid index should not be null.");
                }

                // Test that accessing index 0 (invalid) throws an exception.
                Assert.ThrowsException<ArgumentOutOfRangeException>(() =>
                {
                    var sec = sections[0];
                }, "Accessing index 0 should throw an ArgumentOutOfRangeException.");

                // Test that accessing an index greater than Count throws an exception.
                Assert.ThrowsException<ArgumentOutOfRangeException>(() =>
                {
                    var sec = sections[sections.Count + 1];
                }, "Accessing an index greater than Count should throw an ArgumentOutOfRangeException.");
            });
        }

        [TestMethod]
        public void CKSections_Enumeration_YieldsAllSections()
        {
            LiveWordDocument.WithTestDocument(LiveWordDocument.Default, doc =>
            {
                CKRange range = new CKRange(doc.Content);
                CKSections sections = range.Sections;

                int enumeratedCount = 0;
                foreach (CKSection sec in sections)
                {
                    enumeratedCount++;
                }

                Assert.AreEqual(sections.Count, enumeratedCount,
                    "Enumeration of CKSections should yield the same number of items as the Count property.");
            });
        }

        [TestMethod]
        public void CKSections_ToString_ReturnsValidString()
        {
            LiveWordDocument.WithTestDocument(LiveWordDocument.Default, doc =>
            {
                CKRange range = new CKRange(doc.Content);
                CKSections sections = range.Sections;

                string sectionString = sections.ToString();
                Assert.IsTrue(sectionString.Contains("CKSections"), "ToString() should contain 'CKSections'.");
                Assert.IsTrue(sectionString.Contains("Count:"), "ToString() should contain 'Count:'.");
                Assert.IsTrue(sectionString.Contains(sections.Count.ToString()), "ToString() should include the count value.");
            });
        }
    }
}
