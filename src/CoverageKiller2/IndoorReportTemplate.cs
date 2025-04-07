using CoverageKiller2.DOM;
using Serilog;
using System.IO;
using System.Reflection;
using Word = Microsoft.Office.Interop.Word;

namespace CoverageKiller2
{
    /// <summary>
    /// Represents an indoor report template for PCTEL reports.
    /// It inherits from <see cref="CKDocument"/> and provides functionality
    /// to load the report template from an embedded resource.
    /// </summary>
    [System.Obsolete]
    public class IndoorReportTemplate : CKDocument
    {
        /// <summary>
        /// The name of the embedded resource containing the PCTEL report template.
        /// </summary>
        public static string ResourceName = $"{nameof(CoverageKiller2)}.PCTELReportHeaderFooterTemplate.docx";

        /// <summary>
        /// Initializes a new instance of the <see cref="IndoorReportTemplate"/> class using the specified document path.
        /// </summary>
        /// <param name="wDoc">The path to the Word document.</param>
        //private IndoorReportTemplate(string wDoc) : base(wDoc) { }
        private IndoorReportTemplate(Word.Document wDoc) : base(wDoc) { }

        /// <summary>
        /// Closes the indoor report template without saving changes.
        /// </summary>
        /// <param name="saveChanges">Whether to save changes; this is ignored and always set to false.</param>
        //public override void Close(bool saveChanges = false)
        //{
        //    Log.Information("Closing Indoor Report Template");
        //    base.Close(false);  // Always close without saving changes.
        //}

        /// <summary>
        /// Opens the PCTEL report template from an embedded resource, saves it as a temporary file,
        /// and returns an instance of <see cref="IndoorReportTemplate"/>.
        /// </summary>
        /// <returns>A new instance of <see cref="IndoorReportTemplate"/>.</returns>
        /// 

        public static IndoorReportTemplate OpenResource()
        {
            Log.Information("Opening PCTELDoc template resource...");
            string temporaryFileName = LoadResourceAndCreateTempFile();
            return new IndoorReportTemplate(null);
            //return new IndoorReportTemplate(temporaryFileName);
        }

        /// <summary>
        /// Loads the embedded resource containing the PCTEL report template and writes it to a temporary file.
        /// </summary>
        /// <returns>The path to the temporary file containing the report template.</returns>
        /// <remarks>
        /// This method leverages embedded resources to provide a Word document as a temporary file.
        /// </remarks>
        private static string LoadResourceAndCreateTempFile()
        {
            Log.Debug("Loading resource and creating temporary file...");

            // Get the assembly containing the current class.
            Assembly myAssembly = typeof(IndoorReportTemplate).Assembly;

            // Create a temporary file to store the document.
            string temporaryFileName = Path.GetTempFileName();

            // Load the embedded resource and copy it to the temporary file.
            using (MemoryStream ms = new MemoryStream())
            {
                myAssembly.GetManifestResourceStream(ResourceName)?.CopyTo(ms);
                File.WriteAllBytes(temporaryFileName, ms.ToArray());
            }

            Log.Debug("Resource loaded and temporary file created at {temporaryFileName}", temporaryFileName);

            // Return the filename of the temporary document.
            return temporaryFileName;
        }



    }
}
