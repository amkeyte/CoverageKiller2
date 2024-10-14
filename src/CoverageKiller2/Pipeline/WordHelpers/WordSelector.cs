using Word = Microsoft.Office.Interop.Word;

namespace CoverageKiller2.Pipeline.WordHelpers
{
    internal static class WordSelector
    {
        /// <summary>
        /// Set the Word application selection to skDoc's header.
        /// </summary>
        /// <param name="ckDoc">The document in which to set th selection</param>
        public static void Header(CKDocument ckDoc)
        {
            ckDoc.Activate();
            var activeWindow = ckDoc.WordDoc.ActiveWindow;

            //set up the active window for correct selecting
            if (activeWindow.View.SplitSpecial != Word.WdSpecialPane.wdPaneNone)
            {
                activeWindow.Panes[2].Close();
            }

            if (activeWindow.ActivePane.View.Type == Word.WdViewType.wdNormalView
                || activeWindow.ActivePane.View.Type == Word.WdViewType.wdOutlineView)
            {
                activeWindow.ActivePane.View.Type = Word.WdViewType.wdPrintView;
            }

            //in the active window, go to the header
            activeWindow.ActivePane.View.SeekView = Word.WdSeekView.wdSeekCurrentPageHeader;

            //commit the selection
            ckDoc.WordApp.Selection.WholeStory();
        }

        /// <summary>
        /// Set the Word application selection to skDoc's footer.
        /// </summary>
        /// <param name="ckDoc">The document in which to set th selection</param>
        public static void Footer(CKDocument ckDoc)
        {
            ckDoc.Activate();
            var activeWindow = ckDoc.WordDoc.ActiveWindow;

            //set up the active window for correct selecting
            if (activeWindow.View.SplitSpecial != Word.WdSpecialPane.wdPaneNone)
            {
                activeWindow.Panes[2].Close();
            }

            if (activeWindow.ActivePane.View.Type == Word.WdViewType.wdNormalView
                || activeWindow.ActivePane.View.Type == Word.WdViewType.wdOutlineView)
            {
                activeWindow.ActivePane.View.Type = Word.WdViewType.wdPrintView;
            }

            //in the active window, go to the header
            activeWindow.ActivePane.View.SeekView = Word.WdSeekView.wdSeekCurrentPageFooter;

            //commit the selection
            ckDoc.WordApp.Selection.WholeStory();
        }

        internal static void MainDocument(CKDocument ckDoc)
        {
            ckDoc.Activate();
            var activeWindow = ckDoc.WordDoc.ActiveWindow;

            activeWindow.ActivePane.View.SeekView = Word.WdSeekView.wdSeekMainDocument;
            ckDoc.WordApp.Selection.WholeStory();
        }
    }
}
