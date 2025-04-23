//using Microsoft.Office.Interop.Word;
using CoverageKiller2.DOM;
using CoverageKiller2.Logging;
using CoverageKiller2.Pipeline;
using CoverageKiller2.Pipeline.Config;
using CoverageKiller2.Pipeline.Processes;
using Serilog;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;

namespace CoverageKiller2

{
    /// <summary>
    /// depreciate
    /// </summary>
    public class CkDocHelpers
    {

        /// <summary>
        /// depreciate
        /// </summary>
        public static void FixHeadersAndFooters(Word.Document wDoc)
        {
            var ckDoc = new CKDocument(wDoc);
            var template = IndoorReportTemplate.OpenResource(ckDoc.Application);
            Log.Information("Running Pipeline: Fix Headers and footers for document {Document}", wDoc.Name);
            var pipeline = new CKWordPipeline(ckDoc)
            {
                //{ new GetUserState() },
                //{ new PageSetupFixer(template)  },
                //{ new HeaderFixer(template) },
                { new FooterHeaderFixer(template) },

                //{ new ResetUserState() },
            };
            pipeline.Run();
            Log.Information("Pipeline completed.");

            ckDoc.Application.CloseDocument(template);
            Log.Information("Cleaning up.");
            ckDoc.Activate();
            ckDoc.ActiveWindow.ActivePane.View.SeekView = Word.WdSeekView.wdSeekMainDocument;
        }

        /// <summary>
        /// depreciate
        /// </summary>
        public static void FixPRMCEDoc800(Word.Document wDoc)
        {
            //var ckDoc = new CKDocument(wDoc);
            //var template = IndoorReportTemplate.OpenResource();
            //Log.Information("Running Pipeline: Fix Headers and footers for document {Document}", wDoc.Name);
            //var pipeline = new CKWordPipeline(ckDoc)
            //{
            //    //{ new GetUserState() },
            //    //{ new PageSetupFixer(template)  },
            //    //{ new HeaderFixer(template) },
            //    { new FooterHeaderFixer() },
            //    { new PRMCEFixer("800MHz") },
            //    //{ new ResetUserState() },
            //};
            //pipeline.Run();
            //Log.Information("Pipeline completed.");

            //template.Close();
            //Log.Information("Cleaning up.");
            //ckDoc.Activate();
            //ckDoc.COMObject.ActiveWindow.ActivePane.View.SeekView = Word.WdSeekView.wdSeekMainDocument;
        }
        public static void FixPRMCEDocUHF(Word.Document wDoc)
        {
            //var ckDoc = new CKDocument(wDoc);
            //var template = IndoorReportTemplate.OpenResource();
            //Log.Information("Running Pipeline: Fix Headers and footers for document {Document}", wDoc.Name);
            //var pipeline = new CKWordPipeline(ckDoc)
            //{
            //    //{ new GetUserState() },
            //    //{ new PageSetupFixer(template)  },
            //    //{ new HeaderFixer(template) },
            //    { new FooterHeaderFixer(template) },
            //    { new PRMCEFixer("UHF") },
            //    //{ new ResetUserState() },
            //};
            //pipeline.Run();
            //Log.Information("Pipeline completed.");

            //template.Close();
            //Log.Information("Cleaning up.");
            //ckDoc.Activate();
            //ckDoc.COMObject.ActiveWindow.ActivePane.View.SeekView = Word.WdSeekView.wdSeekMainDocument;
        }

        public static void TestProcessor(CKDocument document)
        {
            LH.Ping<CkDocHelpers>();

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

                foreach (var step in _loader.ProcessorConfig.Pipeline.Steps.StepList)
                {
                    Log.Information($"Step: {step.Name}");
                }

                //var ckDoc = new CKDocument(document);
                var template = IndoorReportTemplate.OpenResource(document.Application);
                Dictionary<string, object> initVars = new Dictionary<string, object>
                    {
                        { nameof(document), document },
                        { nameof(template), template },
                        { nameof(_loader.ProcessorConfig), _loader.ProcessorConfig }
                    };

                var pipeline = new CKWordPipeline(initVars);
                foreach (var x in _loader.ProcessorConfig.Pipeline.Steps.StepList)
                {
                    Log.Debug("var x in _loader...");
                    // Get the step class type dynamically using reflection
                    Type stepType = Type.GetType($"{x.Namespace}.{x.Name}");

                    if (stepType == null)
                    {
                        Log.Error($"Step type '{x.Name}' not found.");
                        continue;
                    }

                    try
                    {
                        Log.Debug("trying...");
                        // Assuming the constructor takes an instance of IndoorReportTemplate
                        CKWordPipelineProcess instance =
                            (CKWordPipelineProcess)Activator.CreateInstance(stepType);
                        pipeline.Add(instance);
                        Log.Information($"Successfully created instance of {x.Name}");
                    }
                    catch (Exception ex)
                    {
                        Log.Error($"Error creating instance of {x.Name}: {ex.Message}");


                        if (Debugger.IsAttached) Debugger.Break();
                    }
                }
                pipeline.Run();
                Log.Information("Pipeline completed.");

                document.Application.CloseDocument(template);


            }
            else
            {
                Log.Error("Failed to load processor configuration.");
            }
            LH.Pong<CkDocHelpers>();
        }


        private static string SelectConfigFile()
        {
            string lastFolder = Properties.Settings.Default.LastUsedFolder;
            string lastFile = Properties.Settings.Default.LastOpenedFile;

            using (OpenFileDialog openFileDialog = new OpenFileDialog())
            {
                openFileDialog.Title = "Select Processor Config File";
                openFileDialog.Filter = "XML Files (*.xml)|*.xml|All Files (*.*)|*.*";

                // Use last opened file if available, otherwise default to last used folder
                if (!string.IsNullOrEmpty(lastFile) && File.Exists(lastFile))
                {
                    openFileDialog.InitialDirectory = Path.GetDirectoryName(lastFile);
                    openFileDialog.FileName = lastFile; // Pre-selects the last file
                }
                else
                {
                    openFileDialog.InitialDirectory = Directory.Exists(lastFolder) ? lastFolder
                                                     : Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
                }

                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    // Save selected folder and file path for next time
                    Properties.Settings.Default.LastUsedFolder = Path.GetDirectoryName(openFileDialog.FileName);
                    Properties.Settings.Default.LastOpenedFile = openFileDialog.FileName;
                    Properties.Settings.Default.Save();

                    return openFileDialog.FileName;
                }
            }

            return null; // User canceled
        }

    }
}