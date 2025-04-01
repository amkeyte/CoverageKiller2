using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace CoverageKiller2.DOM
{
    [TestClass]
    public class CKDocumentsTests
    {
        [TestMethod]
        public void GetByName_ShouldReturnSameInstance()
        {
            LiveWordDocument.WithTestDocument(LiveWordDocument.DefaultTestFile, doc =>
            {
                var expected = CKDocuments.GetByCOMDocument(doc); // ensures it's added
                var actual = CKDocuments.GetByName(doc.FullName);

                Assert.IsNotNull(actual);
                Assert.AreSame(expected, actual);
                Assert.AreEqual(expected.FullPath, actual.FullPath);
            });
        }

        [TestMethod]
        public void GetByName_ShouldReturnNullIfNotAdded()
        {
            var nonexistent = CKDocuments.GetByName("C:\\fake\\nonexistent.docx");
            Assert.IsNull(nonexistent);
        }
    }
}
