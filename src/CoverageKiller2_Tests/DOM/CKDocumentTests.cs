using CoverageKiller2.Tests;     // Contains VstoTestDocumentLoader.
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace CoverageKiller2.DOM
{
    [TestClass]
    public class CKDocumentTests
    {
        // Update the path below to the location of your test document.


        [TestMethod]
        public void CKDocument_Constructor_LoadsDocumentSuccessfully()
        {
            LiveWordDocument.WithTestDocument(doc =>
            {
                // Assert that the document was loaded and processed correctly.
                Assert.IsNotNull(doc, "CKDocument should not be null.");
                Assert.IsTrue(doc.COMDocument.Paragraphs.Count > 0, "The document should contain paragraphs.");
            });
        }

        [TestMethod]
        public void CKDocument_Properties_ReturnExpectedValues()
        {
            LiveWordDocument.WithTestDocument(doc =>
            {
                // CKDocument is instanced in setup.

                // Test that the COMObject property is not null.
                Assert.IsNotNull(doc.COMDocument, "COMObject should not be null.");

                // Test that the WordApp property returns the Application from the COMObject.
                Assert.IsNotNull(doc.Application, "Application should not be null.");
                Assert.AreEqual(doc.COMDocument.Application, doc.Application, "WordApp should match COMObject.Application.");

                // Test that the FullPath property matches the provided document path.
                Assert.AreEqual(LiveWordDocument.DefaultTestFile, doc.FullPath, "FullPath should match the test document path.");
            });
        }


    }
}
