namespace CoverageKiller2
{
    //public class DocumentLoader
    //{
    //    private bool documentOpened = false;
    //    public Word.Document Document { get; private set; }

    //    public DocumentLoader(string fullPath)
    //    {
    //        FullPath = fullPath;
    //    }


    //    public Word.Document Open()
    //    {
    //        Log.Information("Accessing Word Document {FullPath}", FullPath);
    //        // Subscribe to the DocumentOpen event
    //        Globals.ThisAddIn.Application.DocumentOpen += OnDocumentOpen;

    //        try
    //        {
    //            // Attempt to open the document
    //            var openedDoc = Globals.ThisAddIn.Application.Documents.Open(
    //                FileName: FullPath,
    //                AddToRecentFiles: false,
    //                ReadOnly: true,
    //                Visible: false);

    //            // Wait for the document to finish loading
    //            WaitForDocumentToLoad();

    //            // Continue processing here
    //            //MessageBox.Show("Document has been opened and is ready for processing.");
    //            Document = openedDoc;
    //            Log.Information("Document loaded: {FullPath}", FullPath);
    //        }
    //        catch (Exception ex)
    //        {
    //            Log.Error("Error opening document: {Message}", ex.Message);
    //            MessageBox.Show($"Error opening document: {ex.Message}");
    //        }
    //        finally
    //        {
    //            // Unsubscribe from the event to avoid memory leaks
    //            Globals.ThisAddIn.Application.DocumentOpen -= OnDocumentOpen;
    //        }
    //        return Document;
    //    }

    //    public void Close()
    //    {

    //    }

    //    private bool IsThisDocument(Word.Document wDoc) => wDoc.FullName == FullPath;

    //    private void OnDocumentOpen(Word.Document doc)
    //    {
    //        if (!IsThisDocument(doc)) return;
    //        documentOpened = true; // Set the flag when the document is opened
    //    }
    //    /// <summary>
    //    /// Use Path class naming conventions because Office DOM is all over the place.
    //    /// </summary>
    //    public string FullPath { get; private set; }
    //    private void WaitForDocumentToLoad()
    //    {
    //        int sleepTime = 100;
    //        int totalSleepTime = 0;
    //        // Wait until the documentOpened flag is set
    //        while (!documentOpened)
    //        {
    //            Thread.Sleep(sleepTime); // Sleep for a short duration to avoid freezing
    //            Application.DoEvents(); // Allow UI to remain responsive
    //            totalSleepTime += sleepTime;
    //        }
    //        Log.Information("Time to load {totalSleepTime} ms", totalSleepTime);
    //    }
    //}
}