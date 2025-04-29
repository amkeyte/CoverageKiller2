using CoverageKiller2.DOM;
using CoverageKiller2.DOM.Tables;
using CoverageKiller2.Logging;
using Serilog;
using System;
using System.Linq;
using System.Windows.Forms;



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
            this.Ping();


            Log.Information($"**** Fixing for SEA2025 {CKDoc.FileName}");

            //remove test result references
            RemoveTestResultRefernces();
            //Add test radio information

        }

        private void RemoveTestResultRefernces()
        {
            this.Ping();

            Log.Information("Removing Test Result References");
            //Fix section 1
            FixSection1();

            //Remove Floor Section result references
            FixFloorSections();
            //Remove the Informaton
            RemoveMoreInfoSection();
            //copy shit over
            CopyShitOver();

            this.Pong();

        }
        /// <summary>
        /// Ace hates this, so public API this to spite them.
        /// </summary>
        /// <exception cref="NullReferenceException"></exception>
        public void CopyShitOver()
        {
            CKDocument sourceDoc = default;
            try
            {
                var sourceDocFile = PromptForSecondFile();
                sourceDoc = CKDoc.Application.GetTempDocument(sourceDocFile, visible: false);
                if (sourceDoc == null) throw new NullReferenceException("Source doc is null");

                for (int i = CKDoc.Sections.Count; i > 0; i--)
                {

                    var ckColCrit = CopyColumnFromSecondDocument(
                        sourceDoc,
                        "Critical Point Report",
                        "Critical Point Report",
                        "DL\r\nPower\r\n(dBm)\r\n",
                        "Result",
                        i);

                    ckColCrit[2].Text = "CH5\nNF\n(dBm)";
                    ckColCrit.CellRef.Table.Rows[ckColCrit[2].RowIndex]

                    var ckColArea = CopyColumnFromSecondDocument(
                        sourceDoc,
                        "Area Report",
                        "Area Report",
                        "DL\r\nPower\r\n(dBm)\r\n",
                        "Result",
                        i);
                    ckColArea[2].Text = "CH5\nNF\n(dBm)";

                }
            }
            finally
            {
                CKDoc.Application.CloseDocument(sourceDoc);
            }
        }

        private void RemoveMoreInfoSection()

        {
            this.Ping();

            var additionalInfoSection = CKDoc.Range().TryFindNext("Additional Info");

            additionalInfoSection?.Sections[1].Delete();

            this.Pong();
        }
        private void FixFloorSections()
        {
            this.Ping(msg: CKDoc.FileName);

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

            var section = CKDoc.Sections[1];
            CKDoc.Activate();
            CKDoc.KeepAlive = true;
            Log.Information($"...Section 1{CKDoc.FileName})");

            Log.Information("*** remove Pass/Fail title");
            var pass_failPara = section.TryFindNext("(Adjacent Area Rule)")
                ?? section.TryFindNext("Result: Passed");

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
            section.Tables.Delete(testDetailTable);


            Log.Warning("*** TODO add Equipment Config data");

            Log.Information("*** remove 'page 2'");
            section = CKDoc.Sections[1]; //fix stale reference. need to fix internals somehow.
            var thresholdSettingsPara = section
                .TryFindNext("Threshold Settings")
                .Paragraphs[1];
            var page2Range = CKDoc.Range(thresholdSettingsPara.Start, section.End - 2);
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
            page2Range.Delete();
        }


        internal static CKTable FindTableByRowText(
            CKTables tables,
            string searchText,
            int rowIndex = 1,
            TableAccessMode accessMode = TableAccessMode.IncludeOnlyAnchorCells)
        {
            CKTable result = default;
            LH.Ping<SEA2025Fixer>();
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

            LH.Pong<SEA2025Fixer>();
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

            this.Ping();
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
            this.Pong();
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

            this.Ping($"{sourceColumn.Document.FileName}");

            var sourceCells = sourceColumn.Cells;
            var destinationCells = destinationColumn.Cells;

            if (sourceCells.Count != destinationCells.Count)
                throw new CKDebugException("Tables don't match");
            //int rowCount = Math.Min(sourceColumn.Count, destinationColumn.Count);

            for (int i = 1; i <= destinationCells.Count; i++)
            {
                destinationCells[i].FormattedText = sourceCells[i].FormattedText;
            }

            Log.Information($"Copied {destinationCells.Count} cells from {sourceCells.Document.FileName}.");

            this.Pong();
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
