using Serilog;
using System;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Office = Microsoft.Office.Core;
using Word = Microsoft.Office.Interop.Word;

// TODO:  Follow these steps to enable the Ribbon (XML) item:

// 1: Copy the following code block into the ThisAddin, ThisWorkbook, or ThisDocument class.

//  protected override Microsoft.Office.Core.IRibbonExtensibility CreateRibbonExtensibilityObject()
//  {
//      return new CKRibbon();
//  }

// 2. Create callback methods in the "Ribbon Callbacks" region of this class to handle user
//    actions, such as clicking a button. Note: if you have exported this Ribbon from the Ribbon designer,
//    move your code from the event handlers to the callback methods and modify the code to work with the
//    Ribbon extensibility (RibbonX) programming model.

// 3. Assign attributes to the control tags in the Ribbon XML file to identify the appropriate callback methods in your code.  

// For more information, see the Ribbon XML documentation in the Visual Studio Tools for Office Help.


namespace CoverageKiller2
{
    [ComVisible(true)]
    public class CKRibbon : Office.IRibbonExtensibility
    {
        private Office.IRibbonUI ribbon;

        public CKRibbon()
        {
        }

        #region IRibbonExtensibility Members

        public string GetCustomUI(string ribbonID)
        {
            return GetResourceText("CoverageKiller2.CKRibbon.xml");
        }

        #endregion

        #region Ribbon Callbacks
        //Create callback methods here. For more information about adding callback methods, visit https://go.microsoft.com/fwlink/?LinkID=271226

        public void Ribbon_Load(Office.IRibbonUI ribbonUI)
        {
            this.ribbon = ribbonUI;
        }

        public void OnFixPCTELDocButton(Office.IRibbonControl control)
        {
            Log.Information($"Fixing PCTelDoc");
            try
            {

                if (Globals.ThisAddIn.Application.Documents.Count > 0)
                {
                    Word.Document wDoc = Globals.ThisAddIn.Application.ActiveDocument;
                    //CkDocHelpers.FixDasReport(doc);
                    CkDocHelpers.FixHeadersAndFooters(wDoc);
                }
                else
                {
                    Log.Information("This was not a PCTELDoc report. Trying again...");
                    MessageBox.Show("Open a PCTEL Report document.");
                }
            }
            catch (Exception ex)
            {
                Log.Error("{ex}", ex);
                throw ex;

            }
            Log.Information("Done Fixing.");
            Log.Debug("Long wait starts here");
        }

        public void OnFix_PRMCE_PCTELDoc800Button(Office.IRibbonControl control)
        {
            Log.Information($"Fixing PCTelDoc for PRMCE 800");
            try
            {

                if (Globals.ThisAddIn.Application.Documents.Count > 0)
                {
                    Word.Document wDoc = Globals.ThisAddIn.Application.ActiveDocument;
                    //CkDocHelpers.FixDasReport(doc);
                    CkDocHelpers.FixPRMCEDoc800(wDoc);
                }
                else
                {
                    Log.Information("This was not a PCTELDoc report. Trying again...");
                    MessageBox.Show("Open a PCTEL Report document.");
                }
            }
            catch (Exception ex)
            {
                Log.Error("{ex}", ex);
                throw ex;

            }
            Log.Information("Done Fixing.");
            //Log.Debug("Long wait starts here");
        }
        #endregion

        #region Helpers

        private static string GetResourceText(string resourceName)
        {
            Assembly asm = Assembly.GetExecutingAssembly();
            string[] resourceNames = asm.GetManifestResourceNames();
            for (int i = 0; i < resourceNames.Length; ++i)
            {
                if (string.Compare(resourceName, resourceNames[i], StringComparison.OrdinalIgnoreCase) == 0)
                {
                    using (StreamReader resourceReader = new StreamReader(asm.GetManifestResourceStream(resourceNames[i])))
                    {
                        if (resourceReader != null)
                        {
                            return resourceReader.ReadToEnd();
                        }
                    }
                }
            }
            return null;
        }

        #endregion
    }
}
