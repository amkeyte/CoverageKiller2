using CoverageKiller2.DOM;
using System;
using Word = Microsoft.Office.Interop.Word;
namespace CoverageKiller2.Tests
{
    public class LiveWordDocument : IDisposable
    {
        public const string DefaultTestFile = "C:\\Users\\akeyte.PCM\\source\\repos\\CoverageKiller2\\src\\CoverageKiller2_Tests\\TestFiles\\SEA Garage (Noise Floor)_20250313_152027.docx";
        public CKDocument Document { get; private set; }
        public Word.Document WordDocument => Document.COMDocument;
        public Word.Application Application => Document.Application;

        public string FullPath { get; private set; }

        public LiveWordDocument(string fullPath = null)
        {
            FullPath = fullPath ?? DefaultTestFile;
            Document = new CKDocument(FullPath, true);
        }


        public void Close(bool saveChanges = false)
        {
            Document.Close(saveChanges);
        }

        public static void WithTestDocument(Action<CKDocument> testAction)
        {
            using (var loader = new LiveWordDocument(DefaultTestFile))
            {
                try
                {
                    testAction(loader.Document);
                }
                finally
                {
                    loader.Close();
                }
            }
        }

        public static void WithTestDocument(string documentPath, Action<Word.Document> testAction)
        {
            using (var loader = new LiveWordDocument(documentPath))
            {
                try
                {
                    testAction(loader.WordDocument);
                }
                finally
                {
                    loader.Close();
                }
            }
        }

        public void Dispose()
        {
            //anything goes here? CKDocument is already disposable
        }
    }
}
