using CoverageKiller2.Tests; // Contains LiveWordDocument helper.
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using Word = Microsoft.Office.Interop.Word;

namespace CoverageKiller2.DOM
{
    [TestClass]
    public class CKSectionTests
    {
        [TestMethod]
        public void CKSection_Constructor_ThrowsArgumentNullException()
        {
            // Test that passing a null Word.Paragraph throws an ArgumentNullException.
            Assert.ThrowsException<ArgumentNullException>(() =>
            {
                var para = new CKSection(null);
            });
        }

        [TestMethod]
        public void CKSection_Constructor_LoadsSectionSuccessfully()
        {
            LiveWordDocument.WithTestDocument(LiveWordDocument.DefaultTestFile, doc =>
            {
                // Get the first section (Word collections are 1-based).
                Word.Section wordSection = doc.Sections[1];
                CKSection ckSection = new CKSection(wordSection);
                Assert.IsNotNull(ckSection, "CKSection instance should not be null.");
                Assert.IsNotNull(ckSection.COMSection, "COMSection property should not be null.");
            });
        }

        [TestMethod]
        public void CKSection_HeaderRange_GetAndSet_Works()
        {
            LiveWordDocument.WithTestDocument(LiveWordDocument.DefaultTestFile, doc =>
            {
                Word.Section wordSection = doc.Sections[1];
                CKSection ckSection = new CKSection(wordSection);

                // Get the current header range.
                CKRange originalHeader = ckSection.HeaderRange;
                Assert.IsNotNull(originalHeader, "HeaderRange getter should not return null.");

                // For testing, create a new header range using a small portion of the document.
                Word.Range newHeaderWordRange = doc.Range(0, Math.Min(10, doc.Content.Text.Length));
                CKRange newHeader = new CKRange(newHeaderWordRange);

                // Set the header range to the new range.
                ckSection.HeaderRange = newHeader;

                // Verify the header's formatted text was updated.
                CKRange updatedHeader = ckSection.HeaderRange;

                // Normalize text by trimming trailing carriage returns and newlines.
                string expectedText = newHeader.Text.TrimEnd('\r', '\n');
                string actualText = updatedHeader.Text.TrimEnd('\r', '\n');

                Assert.AreEqual(expectedText, actualText, "Header text should match the new header range.");
            });
        }


        [TestMethod]
        public void CKSection_FooterRange_GetAndSet_Works()
        {
            LiveWordDocument.WithTestDocument(LiveWordDocument.DefaultTestFile, doc =>
            {
                Word.Section wordSection = doc.Sections[1];
                CKSection ckSection = new CKSection(wordSection);

                // Get the current footer range.
                CKRange originalFooter = ckSection.FooterRange;
                Assert.IsNotNull(originalFooter, "FooterRange getter should not return null.");

                // For testing, create a new footer range from a portion of the document.
                Word.Range newFooterWordRange = doc.Range(0, Math.Min(10, doc.Content.Text.Length));
                CKRange newFooter = new CKRange(newFooterWordRange);

                // Set the footer range to the new range.
                ckSection.FooterRange = newFooter;

                // Verify the footer's formatted text was updated.
                CKRange updatedFooter = ckSection.FooterRange;

                // Normalize text by trimming trailing carriage returns and newlines.
                string expectedText = newFooter.Text.TrimEnd('\r', '\n');
                string actualText = updatedFooter.Text.TrimEnd('\r', '\n');

                Assert.AreEqual(expectedText, actualText, "Footer text should match the new footer range.");
            });
        }


        [TestMethod]
        public void CKSection_PageSetup_ReturnsValidObject()
        {
            LiveWordDocument.WithTestDocument(LiveWordDocument.DefaultTestFile, doc =>
            {
                Word.Section wordSection = doc.Sections[1];
                CKSection ckSection = new CKSection(wordSection);
                Word.PageSetup pageSetup = ckSection.PageSetup;
                Assert.IsNotNull(pageSetup, "PageSetup property should not be null.");
            });
        }

        [TestMethod]
        public void CKSection_ToString_ReturnsValidString()
        {
            LiveWordDocument.WithTestDocument(LiveWordDocument.DefaultTestFile, doc =>
            {
                Word.Section wordSection = doc.Sections[1];
                CKSection ckSection = new CKSection(wordSection);
                string sectionString = ckSection.ToString();
                Assert.IsTrue(sectionString.Contains("Section:"), "ToString() should include the text 'Section:'.");
                Assert.IsTrue(sectionString.Contains(ckSection.Start.ToString()), "ToString() should include the start position.");
                Assert.IsTrue(sectionString.Contains(ckSection.End.ToString()), "ToString() should include the end position.");
            });
        }
    }
}
