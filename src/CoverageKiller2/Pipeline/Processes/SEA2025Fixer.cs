using CoverageKiller2.DOM;
using CoverageKiller2.DOM.Tables;
using CoverageKiller2.Logging;
using Serilog;
using System;
using System.Linq;
using System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;



namespace CoverageKiller2.Pipeline.Processes
{
    internal class SEA2025Fixer : CKWordPipelineProcess
    {
        private bool _caughtException = false;

        public SEA2025Fixer()
        {
            Tracer.Enabled = true;
        }

        public Tracer Tracer { get; } = new Tracer(typeof(SEA2025Fixer));
        public override void Process()
        {

            Log.Information($"**** Fixing for SEA2025 {CKDoc.FileName}");


            Log.Information("Starting Fixing Section 1");
            //Fix section 1
            FixSection1();
            Log.Information("Finished Fixing Section 1");
            Log.Information("Starting Fixing Floor Sections");
            LongOperationHelpers.TrySilentSave(CKDoc);
            LongOperationHelpers.DoStandardPause();
            //Remove Floor Section result references
            FixFloorSections();
            Log.Information("Done Fixing Floor Sections");
            Log.Information("Removing More Info");
            LongOperationHelpers.TrySilentSave(CKDoc);
            LongOperationHelpers.DoStandardPause();
            //Remove the Informaton
            RemoveMoreInfoSection();
            Log.Information("Done Removing More Info");
            Log.Information("Copying Shit Over");
            LongOperationHelpers.TrySilentSave(CKDoc);
            LongOperationHelpers.DoStandardPause();
            //copy shit over
            CopyShitOver();
            Log.Information("Done Copying shit over.");
            LongOperationHelpers.TrySilentSave(CKDoc);
            LongOperationHelpers.DoStandardPause();

        }


        /// <summary>
        /// Ace hates this, so public API this to spite them.
        /// </summary>
        /// <exception cref="NullReferenceException"></exception>
        public void CopyShitOver()
        {
            CKDocument ch5sourceDoc = default;
            CKDocument iwnSourceDoc = default;
            try
            {
                var ch5sourceDocFile = PromptForSecondFile();
                ch5sourceDoc = CKDoc.Application.GetTempDocument(ch5sourceDocFile, visible: false);
                if (ch5sourceDoc == null) throw new NullReferenceException("Source doc is null");
                Log.Information("Copying Channel 5 data");
                for (int i = CKDoc.Sections.Count; i > 0; i--)
                {

                    var ckColCrit = CopyColumnFromSecondDocument(
                        ch5sourceDoc,
                        "Critical Point Report",
                        "Critical Point Report",
                        "DL\r\nPower\r\n(dBm)\r\n",
                        "Result",
                        i);
                    if (ckColCrit != null)
                    {
                        ckColCrit[2].Text = "CH5\nNF\n(dBm)";
                        ckColCrit.CellRef.Table.Rows[2][ckColCrit.Index + 1].Text = "IWN\nDL Power\n(dBm)";
                    }

                    var ckColArea = CopyColumnFromSecondDocument(
                        ch5sourceDoc,
                        "Area Report",
                        "Area Report",
                        "DL\r\nPower\r\n(dBm)\r\n",
                        "Result",
                        i);
                    if (ckColArea != null)
                    {
                        ckColArea[2].Text = "CH5\nNF\n(dBm)";
                        ckColArea.CellRef.Table.Rows[2][ckColArea.Index + 1].Text = "IWN\nDL Power\n(dBm)";
                    }
                }

                LongOperationHelpers.DoStandardPause();
                var iwnSourceDocFile = PromptForSecondFile();
                iwnSourceDoc = CKDoc.Application.GetTempDocument(iwnSourceDocFile, visible: false);
                if (iwnSourceDoc == null) throw new NullReferenceException("Source doc is null");
                Log.Information("Copying IWN data");
                for (int i = CKDoc.Sections.Count; i > 1; i--) //hacked to 1 to leave column heading
                {
                    var ckColCrit = CopyColumnFromSecondDocument(
                        iwnSourceDoc,
                        "Critical Point Report",
                        "Critical Point Report",
                        "DL\r\nPower\r\n(dBm)\r\n",
                        "IWN\nDL Power\n(dBm)",
                        i);

                    var ckColIWN = CopyColumnFromSecondDocument(
                        iwnSourceDoc,
                        "Area Report",
                        "Area Report",
                        "DL\r\nPower\r\n(dBm)\r\n",
                        "IWN\nDL Power\n(dBm)",
                        i);
                }

            }
            finally
            {
                iwnSourceDoc?.Application.CloseDocument(iwnSourceDoc);
                ch5sourceDoc?.Application.CloseDocument(ch5sourceDoc);
            }
        }

        private void RemoveMoreInfoSection()

        {

            var additionalInfoSection = CKDoc.Range().TryFindNext("Additional Info");

            additionalInfoSection?.Sections[1].Delete();

        }
        private void FixFloorSections()
        {

            foreach (var section in CKDoc.Sections.Reverse())
            {
                Log.Information($"Deleting Floor Section pass/fail subtitle.{CKDoc.FileName}");
                var floorSectionHeadingResult = section.TryFindNext("Result: *", matchWildcards: true);
                floorSectionHeadingResult?.Paragraphs[1]?.Delete();

                Log.Information("*** remove section heading table fields");
                string searchText = "Freq (MHz)\tTech\tBand\tAnt Gain\tCable Loss\tPh.\tType\tMod\tNAC\tArea Points passed (%)\tCritical Points passed (%)";
                var floorSectionHeadingTable = FindTableByRowText(section.Tables, searchText);

                if (floorSectionHeadingTable != null)
                {
                    var headersToRemove = "Ant Gain\tCable Loss\tPh.\tType\tMod\tArea Points passed (%)\tCritical Points passed (%)"
                        .Split(new[] { '\t' }, StringSplitOptions.RemoveEmptyEntries)
                        .Select(s => s.Scrunch());

                    floorSectionHeadingTable.Columns
                        .Delete(col => headersToRemove.Contains(col[1].Text.Scrunch()));


                    floorSectionHeadingTable.MakeFullPage();
                }


                Log.Information("[Issue1]*** remove extra critical point fields");
                searchText = "Critical Point Report";
                var floorSectionCriticalPointsTable = FindTableByRowText(section.Tables,
                    searchText,
                    accessMode: TableAccessMode.IncludeOnlyAnchorCells);//avoid the header cell

                if (floorSectionCriticalPointsTable != null)
                {
                    var headersToRemove = "UL\r\nPower\r\n(dBm)\tUL\r\nS/N\r\n(dB)\tUL\r\nFBER\r\n(%)\r\n"
                        .Split(new[] { '\t' }, StringSplitOptions.RemoveEmptyEntries)
                        .Select(s => s.Scrunch());

                    floorSectionCriticalPointsTable.Columns
                        .Delete(col => headersToRemove.Contains(col[2].Text.Scrunch()));

                    floorSectionCriticalPointsTable.MakeFullPage();
                }

                Log.Information("[Issue1]*** remove extra area point fields");

                searchText = "Area Report";
                var floorSectionAreaReportTable = FindTableByRowText(section.Tables,
                    searchText,
                    accessMode: TableAccessMode.IncludeOnlyAnchorCells);//avoid the header cell
                if (floorSectionAreaReportTable != null)
                {

                    var headersToRemove = "UL\r\nPower\r\n(dBm)\tUL\r\nS/N\r\n(dB)\tUL\r\nFBER\r\n(%)\r\n"
                        .Split(new[] { '\t' }, StringSplitOptions.RemoveEmptyEntries)
                        .Select(s => s.Scrunch());

                    floorSectionAreaReportTable.Columns
                        .Delete(col => headersToRemove.Contains(col[2].Text.Scrunch()));

                    floorSectionAreaReportTable.MakeFullPage();
                }

            }
        }


        private void FixSection1()
        {

            Log.Information("Fixing Header / Footer.");
            Template.CopyHeaderTo(CKDoc);
            Template.CopyFooterTo(CKDoc);

            var section = CKDoc.Sections[1];
            CKDoc.Activate();
            CKDoc.KeepAlive = true;
            Log.Information($"...Section 1{CKDoc.FileName})");

            Log.Information("*** remove Pass/Fail title");
            var pass_failPara = section.TryFindNext("(Adjacent Area Rule)")
                ?? section.TryFindNext("Result: Pass");

            if (pass_failPara?.Paragraphs.Count >= 1)
                pass_failPara.Paragraphs[1].Delete();
            else
                Log.Warning("Pass/Fail title paragraph not found.");


            Log.Information("[Issue 7] *** fix Test Report Summary");
            string searchText = "Channel/ Ch Group\tFreq (MHz)\tTechnology\tBand\tResult\tArea Points\r\npassed (%)\tCritical Points passed (%)\r\n";
            var TRSTable = FindTableByRowText(section.Tables, searchText);

            if (TRSTable != null)
            {
                Log.Debug(TRSTable.Rows.DumpList, "TRSTable Rows");
                var headersToRemove = "Result\tArea Points\r\npassed (%)\tCritical Points passed (%)\r\n"
                    .Split(new[] { '\t' }, StringSplitOptions.RemoveEmptyEntries)
                    .Select(s => s.Scrunch());

                TRSTable.Columns
                    .Delete(col => headersToRemove.Contains(col[1].Text.Scrunch()));

                TRSTable.MakeFullPage();
            }
            else
            {
                Log.Warning("The requested table was not found.");
            }


            Log.Information("*** remove Test Details");

            searchText = "Test Details";
            var testDetailTable = FindTableByRowText(section.Tables, searchText);
            if (testDetailTable != null) section.Tables.Delete(testDetailTable);


            Log.Warning("*** TODO add Equipment Config data");

            Log.Information("*** remove 'page 2'");
            section = CKDoc.Sections[1]; //fix stale reference. need to fix internals somehow.
            var thresholdSettingsPara = section?
                .TryFindNext("Threshold Settings")?
                .Paragraphs[1];
            if (thresholdSettingsPara != null)
            {
                var page2Range = CKDoc.Range(thresholdSettingsPara.Start, section.End - 2);
                page2Range.Delete();

            }

            //page2Range.Delete(page2Range.Tables);

            ////reaquire the same range, since the tables borked up the original
            //thresholdSettingsPara = section
            //     .TryFindNext("Threshold Settings")
            //     .Paragraphs[1];
            //CKDoc.EnsureLayoutReady();
            //var page2Range = CKDoc.Range(thresholdSettingsPara.Start, section.End - 2);
            //CKDoc.EnsureLayoutReady();
            //thresholdSettingsPara.SetBackgroundColor(Microsoft.Office.Interop.Word.WdColor.wdColorRed);
            //CKDoc.EnsureLayoutReady();
            //page2Range.SetBackgroundColor(Microsoft.Office.Interop.Word.WdColor.wdColorBlue);
            //CKDoc.EnsureLayoutReady();
        }


        internal static CKTable FindTableByRowText(
            CKTables tables,
            string searchText,
            int rowIndex = 1,
            TableAccessMode accessMode = TableAccessMode.IncludeOnlyAnchorCells)
        {
            CKTable result = default;
            foreach (var table in tables)
            {

                table.AccessMode = accessMode;
                var rowText = string.Join(string.Empty, table.Rows[rowIndex].Select(c => c.Text));

                LH.Debug($"Searching {LH.GetTableTitle(table, "***Table")} with row {rowIndex} text \n" +
                    $"[{rowText.Scrunch()}] using search text \n[{searchText.Scrunch()}]");

                if (rowText.ScrunchContains(searchText))
                {
                    result = table;
                    LH.Debug("Table found");
                    break;
                }
            }

            return result;
        }


        public CKColumn CopyColumnFromSecondDocument(
           CKDocument sourceDoc,
           string sourceTableSearchText,
           string destinationTableSearchText,
           string sourceHeadingText,
           string destinationHeadingText,
           int sectionIndex)
        {
            if (sourceDoc == null) throw new ArgumentNullException(nameof(sourceDoc));

            CKColumn result = default;

            try
            {
                var sourceTable = FindTableByRowText(sourceDoc.Sections[sectionIndex].Tables, sourceTableSearchText, 1);
                var destinationTable = FindTableByRowText(CKDoc.Sections[sectionIndex].Tables, destinationTableSearchText, 1);

                if (sourceTable == null || destinationTable == null)
                {
                    Log.Warning($"Could not find matching tables for Section {sectionIndex}.");
                    return null;
                }

                var sourceColumn = sourceTable.Columns
                    .FirstOrDefault(col => col[2].Text.ScrunchContains(sourceHeadingText));
                CKColumn destinationColumn = default;
                result = destinationColumn = destinationTable.Columns
                    .FirstOrDefault(col => col[2].Text.ScrunchContains(destinationHeadingText));

                if (sourceColumn == null || destinationColumn == null)
                {
                    Log.Warning($"Source or destination column not found for Section {sectionIndex}.");
                    return null;
                }

                CopyColumn(sourceColumn, destinationColumn);

                if (destinationColumn.Cells.Count >= 2)
                    destinationColumn[2].Text = "Ch. 5 Noise Floor (dBm)";

                Log.Information($"Column copy completed successfully for Section {sectionIndex}.");
            }
            catch (ArgumentOutOfRangeException ex)
            {
                Log.Warning("Section index out of range during column copy.");
                Log.Error(ex.Message);
            }
            return result;
        }

        /// <summary>
        /// Copies text from the source CKColumn to the destination CKColumn.
        /// </summary>
        /// <param name="sourceColumn">The column to copy from.</param>
        /// <param name="destinationColumn">The column to copy into.</param>
        /// <remarks>Version: CK2.00.01.0004</remarks>
        public void CopyColumn(CKColumn sourceColumn, CKColumn destinationColumn)
        {
            if (sourceColumn == null) throw new ArgumentNullException(nameof(sourceColumn));
            if (destinationColumn == null) throw new ArgumentNullException(nameof(destinationColumn));


            var sourceCells = sourceColumn.Cells;
            var destinationCells = destinationColumn.Cells;

            if (sourceCells.Count != destinationCells.Count)
                throw new CKDebugException("Tables don't match");
            //int rowCount = Math.Min(sourceColumn.Count, destinationColumn.Count);

            var templateFont = destinationColumn.CellRef.Table.Rows[3][2].Font;
            for (int i = 1; i <= destinationCells.Count; i++)
            {
                var destinationCell = destinationCells[i];
                destinationCell.Text = sourceCells[i].Text;

                ApplyFontFromTemplate(destinationCell, templateFont);

            }

            Log.Information($"Copied {destinationCells.Count} cells from {sourceCells.Document.FileName}.");

            destinationCells[1].Font.Bold = -1;

            Log.Information($"Set column font.");

        }


        public void ApplyFontFromTemplate(CKCell cell, Word.Font templateFont)
        {
            foreach (var para in cell.Paragraphs)
            {
                var font = para.Font;
                font.Name = templateFont.Name;
                font.Size = templateFont.Size;
                font.Bold = templateFont.Bold;
                font.Italic = templateFont.Italic;
                font.Color = templateFont.Color;
            }
        }
        private string PromptForSecondFile()
        {
            using (var dlg = new OpenFileDialog())
            {
                dlg.RestoreDirectory = true;
                dlg.Title = "Select Source Document";
                dlg.Filter = "Word Documents (*.docx)|*.docx|All Files (*.*)|*.*";
                dlg.Multiselect = false;

                if (dlg.ShowDialog() == DialogResult.OK)
                    return dlg.FileName;

                return null;
            }



        }
    }
}
