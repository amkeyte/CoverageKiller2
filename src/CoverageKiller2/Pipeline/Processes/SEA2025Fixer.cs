using CoverageKiller2.DOM;
using CoverageKiller2.DOM.Tables;
using CoverageKiller2.Logging;
using Serilog;
using System;
using System.Linq;



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


            Log.Information("**** Fixing for SEA2025");

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
            this.Pong();

        }

        private void RemoveMoreInfoSection()

        {
            this.Ping();

            var additionalInfoSection = CKDoc.Range().TryFindNext("Additional Info");
            additionalInfoSection.Sections[1].Delete();

            this.Pong();
        }
        private void FixFloorSections()
        {
            this.Ping();

            foreach (var section in CKDoc.Sections.Reverse())
            {
                Log.Information("Deleting Floor Section pass/fail subtitle.");
                var floorSectionHeadingResult = section.TryFindNext("Result: *", matchWildcards: true);
                floorSectionHeadingResult?.Paragraphs[1]?.Delete();

                Log.Information("*** remove section heading table fields");
                string searchText = "Freq (MHz)\tTech\tBand\tAnt Gain\tCable Loss\tPh.\tType\tMod\tNAC\tArea Points passed (%)\tCritical Points passed (%)";
                var floorSectionHeadingTable = FindTableByRowText(section.Tables, searchText);

                if (floorSectionHeadingTable != null)
                {
                    var headersToRemove = "Ant Gain\tCable Loss\tPh.\tType\tMod\tNAC"
                        .Split(new[] { '\t' }, StringSplitOptions.RemoveEmptyEntries)
                        .Select(s => s.Scrunch());

                    floorSectionHeadingTable.Columns
                        .Delete(col => headersToRemove.Contains(col[1].Text.Scrunch()));


                    floorSectionHeadingTable.MakeFullPage();
                }


                Log.Information("*** remove extra critical point fields");
                searchText = "Critical Point Report";
                var floorSectionCriticalPointsTable = FindTableByRowText(section.Tables,
                    searchText,
                    accessMode: TableAccessMode.ExcludeAllMergedCells);//avoid the header cell

                if (floorSectionHeadingTable != null)
                {
                    var headersToRemove = "UL\r\nPower\r\n(dBm)\tUL\r\nS/N\r\n(dB)\tUL\r\nFBER\r\n(%)\tResult\tDL\r\nLoss\r\n(dB)\r\n"
                        .Split(new[] { '\t' }, StringSplitOptions.RemoveEmptyEntries)
                        .Select(s => s.Scrunch());

                    floorSectionCriticalPointsTable.Columns
                        .Delete(col => headersToRemove.Contains(col[2].Text.Scrunch()));

                    floorSectionCriticalPointsTable.MakeFullPage();
                }

                Log.Information("*** remove extra area point fields");

                searchText = "Area Report";
                var floorSectionAreaReportTable = FindTableByRowText(section.Tables,
                    searchText,
                    accessMode: TableAccessMode.ExcludeAllMergedCells);//avoid the header cell
                if (floorSectionAreaReportTable != null)
                {

                    var headersToRemove = "UL\r\nPower\r\n(dBm)\tUL\r\nS/N\r\n(dB)\tUL\r\nFBER\r\n(%)\tResult\tDL\r\nLoss\r\n(dB)\r\n"
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
            this.Ping();

            var section = CKDoc.Sections[1];

            Log.Information("...Section 1");

            Log.Information("*** remove Pass/Fail title");
            var pass_failPara = section.TryFindNext("(Adjacent Area Rule)")
                ?? CKDoc.Sections[1].TryFindNext("Result: Passed");

            pass_failPara.Paragraphs[1].Delete();

            Log.Information("*** fix Test Report Summary");
            string searchText = "Channel/ Ch Group\tFreq (MHz)\tTechnology\tBand\tResult\tArea Points\r\npassed (%)\tCritical Points passed (%)\r\n";
            var TRSTable = FindTableByRowText(section.Tables, searchText);

            if (TRSTable != null)
            {
                var headersToRemove = "Result\tArea Points\r\npassed (%)\tCritical Points passed (%)\r\n"
                    .Split(new[] { '\t' }, StringSplitOptions.RemoveEmptyEntries)
                    .Select(s => s.Scrunch());

                TRSTable.Columns
                    .Delete(col => headersToRemove.Contains(col[1].Text.Scrunch()));

                TRSTable.MakeFullPage();
            }

            Log.Information("*** remove Test Details");

            searchText = "Test Details";
            var testDetailTable = FindTableByRowText(section.Tables, searchText);
            testDetailTable?.Delete();


            Log.Information("*** TODO add Equipment Config data");


            Log.Information("*** remove 'page 2'");
            var thresholdSettingsPara = section
                .TryFindNext("Threshold Settings")
                .Paragraphs[1];

            var page2Range = CKDoc.Range(thresholdSettingsPara.Start, section.End - 1);
            page2Range.Delete();

            this.Pong();
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
                if (rowText.ScrunchContains(searchText))
                {
                    result = table;
                    break;
                }
            }

            LH.Pong<SEA2025Fixer>();
            return result;
        }

    }
}
