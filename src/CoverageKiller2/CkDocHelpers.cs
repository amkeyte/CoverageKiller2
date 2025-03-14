//using Microsoft.Office.Interop.Word;
using CoverageKiller2.Pipeline;
using CoverageKiller2.Pipeline.Config;
using CoverageKiller2.Pipeline.Processes;
using Microsoft.Office.Interop.Word;
using Serilog;
using System;
using System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;

namespace CoverageKiller2
{
    /// <summary>
    /// depreciate
    /// </summary>
    public class CkDocHelpers
    {
        //public static void FixDasReport(Word.Document doc)
        //{
        //    doc.FixFirstAndLastPages();

        //    List<Word.Section> sections = doc.GetDataPageSections();

        //    foreach (Word.Section section in sections)
        //    {
        //        doc.FixDataPageHeadingText(section, "BDA / DAS Coverage\r");

        //        //Fix tables
        //        foreach (Word.Table table in section.Range.Tables)
        //        {
        //            if (!table.FixMapDataPageTables())
        //            {
        //                table.DelColumnsIf("Critical Point Report",
        //                    "DL Loss (dB)");

        //                table.DelColumnsIf("Area Report",
        //                    "DL Loss (dB)");
        //            }
        //        }
        //    }
        //}

        //public static void FixDlHeadroomReport(Word.Document doc)
        //{
        //    doc.FixFirstAndLastPages();

        //    List<Word.Section> sections = doc.GetDataPageSections();

        //    foreach (Word.Section section in sections)
        //    {
        //        doc.FixDataPageHeadingText(section, "Downlink Dominance Headroom\r");

        //        //Fix tables
        //        foreach (Word.Table table in section.Range.Tables)
        //        {
        //            if (!table.FixMapDataPageTables())
        //            {
        //                table.RenameColumn("DL Power (dBm)", "DL\rHeadroom\r(dBm)\r\a");

        //                table.DelColumnsIf("Critical Point Report",
        //                    "DL DAQ",
        //                    "UL Power (dBm)",
        //                    "UL DAQ",
        //                    "DL Loss (dB)");

        //                table.DelColumnsIf("Area Report",
        //                    "DL DAQ",
        //                    "UL Power (dBm)",
        //                    "UL DAQ",
        //                    "DL Loss (dB)");
        //            }
        //        }
        //    }
        //}


        /// <summary>
        /// depreciate
        /// </summary>
        public static void FixHeadersAndFooters(Word.Document wDoc)
        {
            var ckDoc = new CKDocument(wDoc);
            var template = IndoorReportTemplate.OpenResource();
            Log.Information("Running Pipeline: Fix Headers and footers for document {Document}", wDoc.Name);
            var pipeline = new CKWordPipeline(ckDoc)
            {
                //{ new GetUserState() },
                //{ new PageSetupFixer(template)  },
                { new HeaderFixer(template) },
                { new FooterFixer(template) },

                //{ new ResetUserState() },
            };
            pipeline.Run();
            Log.Information("Pipeline completed.");

            template.Close();
            Log.Information("Cleaning up.");
            ckDoc.Activate();
            ckDoc.COMObject.ActiveWindow.ActivePane.View.SeekView = Word.WdSeekView.wdSeekMainDocument;
        }

        /// <summary>
        /// depreciate
        /// </summary>
        public static void FixPRMCEDoc800(Word.Document wDoc)
        {
            var ckDoc = new CKDocument(wDoc);
            var template = IndoorReportTemplate.OpenResource();
            Log.Information("Running Pipeline: Fix Headers and footers for document {Document}", wDoc.Name);
            var pipeline = new CKWordPipeline(ckDoc)
            {
                //{ new GetUserState() },
                //{ new PageSetupFixer(template)  },
                { new HeaderFixer(template) },
                { new FooterFixer(template) },
                { new PRMCEFixer("800MHz") },
                //{ new ResetUserState() },
            };
            pipeline.Run();
            Log.Information("Pipeline completed.");

            template.Close();
            Log.Information("Cleaning up.");
            ckDoc.Activate();
            ckDoc.COMObject.ActiveWindow.ActivePane.View.SeekView = Word.WdSeekView.wdSeekMainDocument;
        }
        public static void FixPRMCEDocUHF(Word.Document wDoc)
        {
            var ckDoc = new CKDocument(wDoc);
            var template = IndoorReportTemplate.OpenResource();
            Log.Information("Running Pipeline: Fix Headers and footers for document {Document}", wDoc.Name);
            var pipeline = new CKWordPipeline(ckDoc)
            {
                //{ new GetUserState() },
                //{ new PageSetupFixer(template)  },
                { new HeaderFixer(template) },
                { new FooterFixer(template) },
                { new PRMCEFixer("UHF") },
                //{ new ResetUserState() },
            };
            pipeline.Run();
            Log.Information("Pipeline completed.");

            template.Close();
            Log.Information("Cleaning up.");
            ckDoc.Activate();
            ckDoc.COMObject.ActiveWindow.ActivePane.View.SeekView = Word.WdSeekView.wdSeekMainDocument;
        }

        internal static void TestProcessor(Document wDoc)
        {
            try
            {
                string configPath = SelectConfigFile();
                if (string.IsNullOrEmpty(configPath))
                {
                    Log.Warning("No file selected. Aborting.");
                    return;
                }
                var _loader = new ProcessorConfigLoader();
                bool success = _loader.LoadConfig(configPath);

                if (success)
                {
                    Log.Information($"Processor Name: {_loader.ProcessorConfig.Name}");
                    Log.Information($"Description: {_loader.ProcessorConfig.Description}");
                    Log.Information($"Source Template: {_loader.ProcessorConfig.SourceTemplate}");

                    foreach (var step in _loader.ProcessorConfig.PipelineConfig.Steps.StepList)
                    {
                        Log.Information($"Step: {step.Name}");
                    }
                }
                else
                {
                    Log.Error("Failed to load processor configuration.");
                }
            }
            catch (Exception ex)
            {
                Log.Error($"Exception during processor test: {ex.Message}");
            }
        }
        private static string SelectConfigFile()
        {
            using (OpenFileDialog openFileDialog = new OpenFileDialog())
            {
                openFileDialog.Title = "Select Processor Config File";
                openFileDialog.Filter = "XML Files (*.xml)|*.xml|All Files (*.*)|*.*";
                openFileDialog.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);

                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    return openFileDialog.FileName;
                }
            }

            return null; // User canceled
        }
        //public static void FixNativeReport(Word.Document doc)
        //{
        //    doc.FixFirstAndLastPages();

        //    List<Word.Section> sections = doc.GetDataPageSections();

        //    foreach (Word.Section section in sections)
        //    {
        //        doc.FixDataPageHeadingText(section, "System Macro Coverage\r");

        //        //Fix tables
        //        foreach (Word.Table table in section.Range.Tables)
        //        {
        //            if (!table.FixMapDataPageTables())
        //            {
        //                table.DelColumnsIf("Critical Point Report",
        //                    "DL Loss (dB)");

        //                table.DelColumnsIf("Area Report",
        //                    "DL Loss (dB)");
        //            }
        //        }
        //    }
        //}

        //public static void FixUlHeadroomReport(Word.Document doc)
        //{
        //    doc.FixFirstAndLastPages();

        //    List<Word.Section> sections = doc.GetDataPageSections();

        //    foreach (Word.Section section in sections)
        //    {
        //        doc.FixDataPageHeadingText(section, "Uplink Dominance Headroom\r");

        //        //Fix tables
        //        foreach (Word.Table table in section.Range.Tables)
        //        {
        //            if (!table.FixMapDataPageTables())
        //            {
        //                table.RenameColumn("DL Power (dBm)", "UL\rHeadroom\r(dBm)\r\a");

        //                table.DelColumnsIf("Critical Point Report",
        //                    "DL DAQ",
        //                    "UL Power (dBm)",
        //                    "UL DAQ",
        //                    "DL Loss (dB)");

        //                table.DelColumnsIf("Area Report",
        //                    "DL DAQ",
        //                    "UL Power (dBm)",
        //                    "UL DAQ",
        //                    "DL Loss (dB)");
        //            }
        //        }
        //    }
        //}
    }
}