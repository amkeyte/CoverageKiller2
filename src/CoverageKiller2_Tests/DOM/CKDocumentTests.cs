using CoverageKiller2.DOM;
using CoverageKiller2.Test;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.Linq;

namespace CoverageKiller2.DOM
{
    /// <summary>
    /// Tests for the CKDocument class.
    /// </summary>
    [TestClass]
    // Version: CK2.00.00.0002
    public class CKDocumentTests
    {
        //******* Standard Rigging ********
        private string _testFilePath;
        private CKDocument _testFile;

        [TestInitialize]
        public void Setup()
        {
            _testFilePath = RandomTestHarness.TestFile1;
            _testFile = RandomTestHarness.GetTempDocumentFrom(_testFilePath);
        }
        [TestCleanup]
        public void Cleanup()
        {
            RandomTestHarness.CleanUp(_testFile, force: true);
        }
        //******* End Standard Rigging ********



        [TestMethod]
        public void CKDocument_Constructor_LoadsDocumentSuccessfully()
        {

            Assert.IsNotNull(_testFile);
            Assert.IsTrue(_testFile.Range().Text.Length > 0, "Document should contain text.");
        }

        //[TestMethod]
        //public void CKDocument_Properties_ReturnExpectedValues()
        //{
        //    Assert.IsNotNull(_testFile.Application);
        //    Assert.AreEqual(_testFilePath, _testFile.FullPath);
        //}

        [TestMethod]
        public void CKDocument_Range_WrapsEntireDocument()
        {
            var range = _testFile.Range();
            Assert.IsNotNull(range);
            Assert.IsTrue(range.Text.Length > 0);
        }

        [TestMethod]
        public void CKDocument_CopyHeaderAndFooter_CopiesCorrectly()
        {
            var source = _testFile;
            var target = RandomTestHarness.GetTempDocumentFrom(_testFilePath);
            source.CopyHeaderAndFooterTo(target);

            var sourceHeader = source.GetHeaderRange().Text.Trim();
            var targetHeader = target.GetHeaderRange().Text.Trim();
            var sourceFooter = source.GetFooterRange().Text.Trim();
            var targetFooter = target.GetFooterRange().Text.Trim();

            Assert.AreEqual(sourceHeader, targetHeader, "Header content should match.");
            Assert.AreEqual(sourceFooter, targetFooter, "Footer content should match.");
        }

        [TestMethod]
        public void CKDocument_DeleteSection_RemovesSectionSuccessfully()
        {
            var sectionsBefore = _testFile.Sections.Count;
            if (sectionsBefore < 2)
                Assert.Inconclusive("Test requires a document with at least two sections.");

            _testFile.DeleteSection(2);
            var sectionsAfter = _testFile.Sections.Count;

            Assert.AreEqual(sectionsBefore - 1, sectionsAfter);
        }

        [TestMethod]
        public void CKDocument_Tables_CollectionIsAccessible()
        {
            var tables = _testFile.Tables;
            Assert.IsNotNull(tables);
            Assert.IsTrue(tables.Count >= 0);
        }

        [TestMethod]
        public void CKDocument_Sections_CollectionIsAccessible()
        {
            var sections = _testFile.Sections;
            Assert.IsNotNull(sections);
            Assert.IsTrue(sections.Count >= 1);
        }

        [TestMethod]
        public void CKDocument_IsOrphan_ReturnsFalseForActiveDocument()
        {
            Assert.IsFalse(_testFile.IsOrphan);
        }

        [TestMethod]
        public void CKDocument_Dispose_UntracksFromApplication()
        {
            var app = _testFile.Application;
            Assert.IsTrue(app.Documents.Contains(_testFile));
            _testFile.Dispose();
            Assert.IsFalse(app.Documents.Contains(_testFile));
        }
    }
}
