using Microsoft.Office.Core;
using Serilog;
using System;
using System.IO;
using Word = Microsoft.Office.Interop.Word;

namespace CoverageKiller2.Pipeline.Processes
{
    //Take the signal level values from the noise report and insert them into the tables if the CC test report
    internal class MergeTableValues : CKWordPipelineProcess
    {

        public string InsertAfterFeild { get; set; }
        public string ExtractFeild { get; set; }
        public string SrcSyncSectionText = "Floor Garage L1";
        public string TrgtSyncSectionText = "Floor Garage L1";

        public override void Process()
        {
            GetSourceDoc();
            CopyFeilds();
        }

        private void CopyFeilds()
        {
            foreach (var source in SourceDoc.Sections)
            {
                var target = FindMatchingSection(source);

                Log.Debug($"   Source first \"{source.Paragraphs[1].Text}\"; second: \"{source.Paragraphs[2].Text}\"");
                Log.Debug($"   Target first \"{target.Paragraphs[1].Text}\"; second: \"{target.Paragraphs[2].Text}\"");


            }

        }
        /// <summary>
        /// Finds the matching target section in the target document (CKDoc) based on the first paragraph of the source section.
        /// </summary>
        /// <param name="source">The source CKSection whose first paragraph will be used for matching.</param>
        /// <returns>
        /// The matching CKSection from the target document if found; otherwise, null.
        /// </returns>
        /// <exception cref="InvalidOperationException">
        /// Thrown if the source section does not contain at least one paragraph.
        /// </exception>
        public CKSection FindMatchingSection(CKSection source)
        {
            // Ensure the source section has at least one paragraph.
            if (source.Range.Paragraphs.Count < 1)
                throw new InvalidOperationException("Source section must have at least one paragraph.");

            // Get the trimmed text from the first paragraph of the source section.
            string sourcePara1 = source.Range.Paragraphs[1].Text.Trim();

            // Iterate over each section in the target document (assumed available via CKDoc.Sections).
            foreach (CKSection targetSection in CKDoc.Sections)
            {
                var targetParas = targetSection.Range.Paragraphs;
                if (targetParas.Count < 1)
                {
                    // Skip sections with no paragraphs.
                    continue;
                }

                // Get the trimmed text from the first paragraph of the target section.
                string targetPara1 = targetParas[1].Text.Trim();

                // Compare the first paragraphs using an invariant culture comparison.
                if (string.Equals(sourcePara1, targetPara1, StringComparison.InvariantCulture))
                {
                    return targetSection;
                }
            }

            // No matching section was found.
            return null;

        }


        public void ListFirstParagraphTexts()
        {
            if (CKDoc == null) throw new ArgumentNullException(nameof(CKDoc));
            if (SourceDoc == null) throw new ArgumentNullException(nameof(SourceDoc));

            // Use the minimum number of sections if they differ.
            int sectionCount = Math.Min(CKDoc.Sections.Count, SourceDoc.Sections.Count);

            for (int i = 1; i <= sectionCount; i++)
            {
                // Get corresponding sections (using one-based indexing)
                CKSection sectionDoc = CKDoc.Sections[i];
                CKSection sectionSource = SourceDoc.Sections[i];

                // Retrieve the first paragraph from each section.
                // Note: Word.Paragraphs is one-based.
                CKParagraph paraDoc = sectionDoc.Paragraphs[1];
                CKParagraph paraSource = sectionSource.Paragraphs[1];
                var tablesDoc = sectionDoc.Tables;
                var tablesSource = sectionSource.Tables;

                // Output the text (trimmed) from each first paragraph.
                Log.Debug($"Section {i}:");
                Log.Debug($"   Doc first paragraph: \"{paraDoc.Text.Trim()}\"; first table: \"{tablesDoc[1]}\"");
                Log.Debug($"   Source first paragraph: \"{paraSource.Text.Trim()}\"; first table \"{tablesSource[1]}\"");




            }
        }

        public CKDocument SourceDoc { get; set; }
        private void GetSourceDoc()
        {
            // Determine the default folder: if SourceDoc exists and has a FullPath, use its directory;
            // otherwise, fallback to My Documents.
            string defaultFolder = CKDoc != null && !string.IsNullOrEmpty(CKDoc.FullPath)
                ? Path.GetDirectoryName(CKDoc.FullPath)
                : Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);

            // Get the Word application instance (assuming you're in a VSTO add-in)
            Word.Application app = Globals.ThisAddIn.Application;

            // Initialize and configure the file dialog for file picking.
            FileDialog fileDialog = app.FileDialog[MsoFileDialogType.msoFileDialogFilePicker];
            fileDialog.Title = "Select a Word Document";
            fileDialog.InitialFileName = defaultFolder + Path.DirectorySeparatorChar;
            fileDialog.Filters.Clear();
            fileDialog.Filters.Add("Word Documents", "*.doc;*.docx", 1);

            // Display the dialog. The Show method returns -1 if the user clicks OK.
            if (fileDialog.Show() == -1)
            {
                // Get the selected file (first selected item)
                string selectedFile = fileDialog.SelectedItems.Item(1);

                // Create a new CKDocument using the selected file path.
                SourceDoc = new CKDocument(selectedFile);
            }
        }
    }
}
