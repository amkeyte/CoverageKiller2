using CoverageKiller2.DOM;
using CoverageKiller2.Test;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;

namespace CoverageKiller2.Tests.DOM
{
    /// <summary>
    /// Unit tests for <see cref="CKSections"/> collection wrapper.
    /// </summary>
    /// <remarks>
    /// Version: CK2.00.00.0000
    /// </remarks>
    [TestClass]
    public class CKSectionsTests
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
        public void CKSections_Count_MatchesUnderlyingWordSectionsCount()
        {
            CKRange range = _doc.Range();
            CKSections sections = range.Sections;

            int expectedCount = _doc.Range().Sections.Count;
            Assert.AreEqual(expectedCount, sections.Count,
                "CKSections count should match the Word document's Sections count.");
        }

        [TestMethod]
        public void CKSections_Indexer_ReturnsValidCKSection_And_ThrowsOnInvalidIndex()
        {
            CKRange range = _doc.Range();
            CKSections sections = range.Sections;

            if (sections.Count > 0)
            {
                CKSection section1 = sections[1];
                Assert.IsNotNull(section1, "CKSection returned by a valid index should not be null.");
            }

            Assert.ThrowsException<ArgumentOutOfRangeException>(() =>
            {
                var sec = sections[0];
            }, "Accessing index 0 should throw an ArgumentOutOfRangeException.");

            Assert.ThrowsException<ArgumentOutOfRangeException>(() =>
            {
                var sec = sections[sections.Count + 1];
            }, "Accessing an index greater than Count should throw an ArgumentOutOfRangeException.");
        }

        [TestMethod]
        public void CKSections_Enumeration_YieldsAllSections()
        {
            CKRange range = _doc.Range();
            CKSections sections = range.Sections;

            int enumeratedCount = 0;
            foreach (CKSection sec in sections)
            {
                enumeratedCount++;
            }

            Assert.AreEqual(sections.Count, enumeratedCount,
                "Enumeration of CKSections should yield the same number of items as the Count property.");
        }

        [TestMethod]
        public void CKSections_ToString_ReturnsValidString()
        {
            CKRange range = _doc.Range();
            CKSections sections = range.Sections;

            string sectionString = sections.ToString();
            Assert.IsTrue(sectionString.Contains("CKSections"), "ToString() should contain 'CKSections'.");
            Assert.IsTrue(sectionString.Contains("Count:"), "ToString() should contain 'Count:'.");
            Assert.IsTrue(sectionString.Contains(sections.Count.ToString()), "ToString() should include the count value.");
        }
    }
}
