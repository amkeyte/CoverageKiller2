using Serilog;
using System;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Office = Microsoft.Office.Core;
using Word = Microsoft.Office.Interop.Word;

namespace CoverageKiller2
{
    /// <summary>
    /// Custom Ribbon class for the Word add-in, implementing the IRibbonExtensibility interface.
    /// It includes callback methods for various ribbon buttons to fix PCTEL documents.
    /// </summary>
    [ComVisible(true)]
    public class CKRibbon : Office.IRibbonExtensibility
    {
        private Office.IRibbonUI ribbon;

        /// <summary>
        /// Initializes a new instance of the <see cref="CKRibbon"/> class.
        /// </summary>
        public CKRibbon() { }

        #region IRibbonExtensibility Members

        /// <summary>
        /// Loads the custom Ribbon XML.
        /// </summary>
        /// <param name="ribbonID">The Ribbon ID to load.</param>
        /// <returns>The Ribbon XML as a string.</returns>
        public string GetCustomUI(string ribbonID)
        {
            return GetResourceText("CoverageKiller2.CKRibbon.xml");
        }

        #endregion

        #region Ribbon Callbacks

        /// <summary>
        /// Called when the Ribbon is loaded.
        /// </summary>
        /// <param name="ribbonUI">The Ribbon UI instance.</param>
        public void Ribbon_Load(Office.IRibbonUI ribbonUI)
        {
            this.ribbon = ribbonUI;
        }

        /// <summary>
        /// Callback for the "Fix PCTEL Doc" button.
        /// Attempts to fix headers and footers in the active PCTEL document.
        /// </summary>
        /// <param name="control">The Ribbon control that triggered the callback.</param>
        public void OnFixPCTELDocButton(Office.IRibbonControl control)
        {
            Log.Information("Fixing PCTEL Document");
            try
            {
                if (Globals.ThisAddIn.Application.Documents.Count > 0)
                {
                    Word.Document wDoc = Globals.ThisAddIn.Application.ActiveDocument;
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
                Log.Error("An error occurred: {Message}", ex.Message);
                MessageBox.Show($"Error: {ex.Message}");
                throw;
            }

            Log.Information("Done fixing PCTEL Document.");
            Log.Debug("Long wait starts here");
        }

        /// <summary>
        /// Callback for the "Fix PRMCE PCTEL Doc 800" button.
        /// Attempts to fix the PRMCE 800 version of the PCTEL document.
        /// </summary>
        /// <param name="control">The Ribbon control that triggered the callback.</param>
        public void OnFix_PRMCE_PCTELDoc800Button(Office.IRibbonControl control)
        {
            Log.Information("Fixing PRMCE 800 PCTEL Document");
            try
            {
                if (Globals.ThisAddIn.Application.Documents.Count > 0)
                {
                    Word.Document wDoc = Globals.ThisAddIn.Application.ActiveDocument;
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
                Log.Error("An error occurred: {Message}", ex.Message);
                MessageBox.Show($"Error: {ex.Message}");
                throw;
            }

            Log.Information("Done fixing PRMCE 800 PCTEL Document.");
        }

        #endregion

        #region Helpers

        /// <summary>
        /// Retrieves the embedded resource text for the specified resource name.
        /// </summary>
        /// <param name="resourceName">The name of the resource to retrieve.</param>
        /// <returns>The content of the resource as a string.</returns>
        private static string GetResourceText(string resourceName)
        {
            Assembly asm = Assembly.GetExecutingAssembly();
            string[] resourceNames = asm.GetManifestResourceNames();

            foreach (string resource in resourceNames)
            {
                if (string.Equals(resourceName, resource, StringComparison.OrdinalIgnoreCase))
                {
                    using (StreamReader resourceReader = new StreamReader(asm.GetManifestResourceStream(resource)))
                    {
                        return resourceReader?.ReadToEnd();
                    }
                }
            }
            return null;
        }

        #endregion
    }
}
