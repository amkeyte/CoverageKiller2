using Serilog;
using System.IO;
using System.Reflection;

namespace CoverageKiller2
{
    public class IndoorReportTemplate : CKDocument

    {
        //keep up with namespace changes
        public static string ResourceName = $"{nameof(CoverageKiller2)}.PCTELReportHeaderFooterTemplate.docx";

        //for future use and clarity
        public override void Close(bool saveChanges = false)
        {
            Log.Information("Closing Indoor Report Template");
            base.Close(false);
        }
        public static IndoorReportTemplate OpenResource()
        {
            Log.Information("Opening PCTELDoc template resource...");
            string temporaryFileName = LoadResourceAndCreateTempFile();
            return new IndoorReportTemplate(temporaryFileName);

            //Word.Document wDoc = new DocumentLoader(temporaryFileName).Open();
        }


        private IndoorReportTemplate(string wDoc) : base(wDoc)
        {

        }



        private static string LoadResourceAndCreateTempFile()
        {
            //https://stackoverflow.com/questions/4367311/embed-a-word-document-in-c-sharp?rq=3
            //https://stackoverflow.com/questions/33164270/how-to-open-embedded-resource-word-document?rq=3
            //https://stackoverflow.com/questions/15925801/visual-studio-c-sharp-how-to-add-a-doc-file-as-a-resource

            //DebugX();

            //here we load the internal resource and save it to a temporary file that can be accessed by Word

            //use only the dll containig this class
            Assembly myAssembly = typeof(IndoorReportTemplate).Assembly;
            //get and create the temp file to store the document
            string temporaryFileName = Path.GetTempFileName();

            //write the file out to the new temp file.
            MemoryStream ms = new MemoryStream();
            myAssembly.GetManifestResourceStream(ResourceName).CopyTo(ms);
            File.WriteAllBytes(temporaryFileName, ms.ToArray());

            //return the filename for the extracted document.
            return temporaryFileName;
        }


    }
}
