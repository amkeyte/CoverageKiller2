using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.Diagnostics;

namespace CoverageKiller2.DOM.Tables
{
    [TestClass]
    public class ShadowWorkspaceTests
    {
        public int UseTable = 2;

        [TestMethod]
        public void ShadowWorkspace_DebuggerViewStaysOpenIfRequested()
        {
            LiveWordDocument.WithTestDocument(doc =>
            {
                var table = doc.Tables[2].COMTable;
                var shadow = new ShadowWorkspace(doc.Application);

                try
                {
                    // View stays open after test run if true
                    shadow.ShowDebuggerWindow(keepOpen: true);

                    var clone = shadow.CloneTable(table);
                    var grid = TableGridCrawler3.NormalizeVisualGrid(clone);

                    Debug.WriteLine("=== Shadow Grid Dump ===");
                    Debug.WriteLine(TableGridCrawler3.DumpGrid(grid));
                }
                finally
                {
                    shadow.Dispose(); // will skip cleanup if keepOpen = true
                }
            });
        }

    }
}
