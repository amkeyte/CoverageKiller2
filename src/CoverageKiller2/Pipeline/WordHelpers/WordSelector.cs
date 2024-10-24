using Word = Microsoft.Office.Interop.Word;

namespace CoverageKiller2.Pipeline.WordHelpers
{
    /// <summary>
    /// Provides methods to set the Word application selection to different sections of a CKDocument.
    /// </summary>
    internal static class WordSelector
    {
        /// <summary>
        /// Sets the Word application selection to the header of the specified document.
        /// </summary>
        /// <param name="ckDoc">The document in which to set the selection to the header.</param>
        public static void Header(CKDocument ckDoc)
        {
            SetSelectionToView(ckDoc, Word.WdSeekView.wdSeekCurrentPageHeader);
        }

        /// <summary>
        /// Sets the Word application selection to the footer of the specified document.
        /// </summary>
        /// <param name="ckDoc">The document in which to set the selection to the footer.</param>
        public static void Footer(CKDocument ckDoc)
        {
            SetSelectionToView(ckDoc, Word.WdSeekView.wdSeekCurrentPageFooter);
        }

        /// <summary>
        /// Sets the Word application selection to the main document of the specified document.
        /// </summary>
        /// <param name="ckDoc">The document in which to set the selection to the main document.</param>
        internal static void MainDocument(CKDocument ckDoc)
        {
            SetSelectionToView(ckDoc, Word.WdSeekView.wdSeekMainDocument);
        }

        /// <summary>
        /// Helper method to set the Word application selection based on the specified view.
        /// </summary>
        /// <param name="ckDoc">The document in which to set the selection.</param>
        /// <param name="seekView">The type of view to set for selection.</param>
        private static void SetSelectionToView(CKDocument ckDoc, Word.WdSeekView seekView)
        {
            ckDoc.Activate();
            var activeWindow = ckDoc.COMObject.ActiveWindow;

            // Ensure the active window is set up correctly for selection
            CloseSplitPaneIfOpen(activeWindow);
            SetViewToPrint(activeWindow);

            // Set the view to the specified seek view
            activeWindow.ActivePane.View.SeekView = seekView;

            // Commit the selection
            ckDoc.WordApp.Selection.WholeStory();
        }

        /// <summary>
        /// Closes the split pane if it is open.
        /// </summary>
        /// <param name="activeWindow">The active window of the document.</param>
        private static void CloseSplitPaneIfOpen(Word.Window activeWindow)
        {
            if (activeWindow.View.SplitSpecial != Word.WdSpecialPane.wdPaneNone)
            {
                activeWindow.Panes[2].Close();
            }
        }

        /// <summary>
        /// Sets the active window view to print view if it is currently in normal or outline view.
        /// </summary>
        /// <param name="activeWindow">The active window of the document.</param>
        private static void SetViewToPrint(Word.Window activeWindow)
        {
            if (activeWindow.ActivePane.View.Type == Word.WdViewType.wdNormalView ||
                activeWindow.ActivePane.View.Type == Word.WdViewType.wdOutlineView)
            {
                activeWindow.ActivePane.View.Type = Word.WdViewType.wdPrintView;
            }
        }
    }
}
