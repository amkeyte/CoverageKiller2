using CoverageKiller2.Test;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;

namespace CoverageKiller2.DOM
{
    /// <summary>
    /// Unit tests for <see cref="CKSection"/> wrapper.
    /// </summary>
    /// <remarks>
    /// Version: CK2.00.00.0000
    /// </remarks>
    [TestClass]
    public class CKSectionTests
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
        public void CKSection_Constructor_ThrowsArgumentNullException()
        {
            Assert.ThrowsException<ArgumentNullException>(() =>
            {
                var para = new CKSection(null, null);
            });
        }

        [TestMethod]
        public void CKSection_Constructor_LoadsSectionSuccessfully()
        {
            var section = _doc.Sections[1];
            Assert.IsNotNull(section, "CKSection instance should not be null.");
            Assert.IsNotNull(section.COMSection, "COMSection property should not be null.");
        }

        [TestMethod]
        public void CKSection_HeaderRange_GetAndSet_Works()
        {
            var section = _doc.Sections[1];
            var originalHeader = section.HeaderRange;
            Assert.IsNotNull(originalHeader, "HeaderRange getter should not return null.");

            var shortRange = _doc.Range(0, Math.Min(10, _doc.Range().Text.Length));
            section.HeaderRange = shortRange;

            var updatedHeader = section.HeaderRange;
            string expectedText = shortRange.Text.TrimEnd('\r', '\n');
            string actualText = updatedHeader.Text.TrimEnd('\r', '\n');
            Assert.AreEqual(expectedText, actualText, "Header text should match the new header range.");
        }

        [TestMethod]
        public void CKSection_FooterRange_GetAndSet_Works()
        {
            var section = _doc.Sections[1];
            var originalFooter = section.FooterRange;
            Assert.IsNotNull(originalFooter, "FooterRange getter should not return null.");

            var shortRange = _doc.Range(0, Math.Min(10, _doc.Range().Text.Length));
            section.FooterRange = shortRange;

            var updatedFooter = section.FooterRange;
            string expectedText = shortRange.Text.TrimEnd('\r', '\n');
            string actualText = updatedFooter.Text.TrimEnd('\r', '\n');
            Assert.AreEqual(expectedText, actualText, "Footer text should match the new footer range.");
        }

        [TestMethod]
        public void CKSection_PageSetup_ReturnsValidObject()
        {
            var section = _doc.Sections[1];
            var pageSetup = section.PageSetup;
            Assert.IsNotNull(pageSetup, "PageSetup property should not be null.");
        }

        [TestMethod]
        public void CKSection_ToString_ReturnsValidString()
        {
            var section = _doc.Sections[1];
            string sectionString = section.ToString();

            Assert.IsTrue(sectionString.Contains("Section:"), "ToString() should include the text 'Section:'.");
            Assert.IsTrue(sectionString.Contains(section.Start.ToString()), "ToString() should include the start position.");
            Assert.IsTrue(sectionString.Contains(section.End.ToString()), "ToString() should include the end position.");
        }
    }
}
