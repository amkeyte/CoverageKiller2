using CoverageKiller2.DOM;
using CoverageKiller2.DOM.Tables;
using CoverageKiller2.Logging;
using CoverageKiller2.Pipeline.WordHelpers;
using Serilog;
using System;
using System.Linq;
using Word = Microsoft.Office.Interop.Word;

///**************************
///
/// I am supressing this file to perform testing on the DOM.
/// 
/// 
/// 
/// 
///**************************



namespace CoverageKiller2.Pipeline.Processes
{
    internal class SEA2025Fixer : CKWordPipelineProcess
    {

        public SEA2025Fixer()
        {
            Tracer.Enabled = true;
        }

        public Tracer Tracer { get; } = new Tracer(typeof(SEA2025Fixer));
        public override void Process()
        {
            Log.Information("**** Fixing for SEA2025");

            //remove test result references
            RemoveTestResultRefernces();
            //Add test radio information
        }

        private void RemoveTestResultRefernces()
        {
            Log.Information("Removing Test Result References");
            //Remove Page 1 result references
            RemoveTestResultReferencesPage1();
            //Remove Page 2
            RemoveTestResultReferencesPage2();
            //Remove Floor Section result references
            RemoveTestResultReferencesFloorSections();
            //Remove the Informaton
            RemoveMoreInfoSection();
        }

        private void RemoveMoreInfoSection()

        {
            var tf = new TextFinder(CKDoc, "Additional Info");
            if (tf.TryFind(out Word.Range foundRange))
            {
                foundRange = FixerHelpers.FindSectionBreak(foundRange, false);
                foundRange.End = foundRange.Document.Content.End;
                foundRange.HighlightColorIndex = Word.WdColorIndex.wdBlue;
                foundRange.Delete();

            }
        }
        private void RemoveTestResultReferencesFloorSections()
        {
            Log.Information("Deleting Floor Section pass/fail subtitle.");
            var tf1 = new TextFinder(CKDoc, "Result: *");
            while (tf1.TryFind(out var fr1, true))
            {
                fr1.Expand(Word.WdUnits.wdParagraph);
                fr1.Delete();
            }

            //Log.Information("*** remove section heading table fields");
            //foreach (var table in CKDoc.Tables
            //    .Where(t => t.RowMatches(1, "Freq (MHz)\tTech\tBand\tAnt Gain\tCable Loss\tPh.\tType\tMod\tNAC\tArea Points passed (%)\tCritical Points passed (%)")))
            //{
            //    FixFloorSectionSectionHeadingTable(table);
            //}
            //Log.Information("*** remove extra critical point fields");
            //foreach (var table in CKDoc.Tables
            //    .Where(t => t.RowMatches(1, "Critical Point Report")))
            //{
            //    FixFloorSectionCriticalPointReportTable(table);
            //}
            //Log.Information("*** remove Area Report point fields");
            //foreach (var table in CKDoc.Tables
            //    .Where(t => t.RowMatches(1, "Area Report")))
            //{
            //    FixFloorSectionAreaReportTable(table);
            //}
        }

        private void FixFloorSectionSectionHeadingTable(CKTable table)
        {
            var headersToRemove = "Area Points passed (%)\tCritical Points passed (%)"
                .Split(new[] { '\t' }, StringSplitOptions.RemoveEmptyEntries)
                .Select(s => FixerHelpers.NormalizeMatchString(s))
                .Reverse()
                .ToList();

            table.Columns
                .Where(col => headersToRemove
                    .Contains(FixerHelpers.NormalizeMatchString(col.Cells[1].Text)))
                .Reverse().ToList().ForEach(col => col.Delete());

            table.MakeFullPage();
        }

        private void FixFloorSectionCriticalPointReportTable(CKTable table)
        {
            Tracer.Log("Entering", "**", new DataPoints($"{nameof(table)}.Index", table.Index));


            Tracer.Log("Deleting first row");
            table.Rows.First().Delete();

            var headersToRemove = "UL\r\nPower\r\n(dBm)\tUL\r\nS/N\r\n(dB)\tUL\r\nFBER\r\n(%)\tResult\tDL\r\nLoss\r\n(dB)\r\n"
                .Split(new[] { '\t' }, StringSplitOptions.RemoveEmptyEntries)
                .Select(s => FixerHelpers.NormalizeMatchString(s))
                .Reverse()
                .ToList();

            table.Columns
                .Where(col => headersToRemove
                    .Contains(FixerHelpers.NormalizeMatchString(col.Cells[1].Text)))
                .Reverse().ToList().ForEach(col => col.Delete());

            table.AddAndMergeFirstRow("Critical Point Report");
            table.MakeFullPage();

        }
        private void FixFloorSectionAreaReportTable(CKTable table)
        {
            Tracer.Log("Entering", "**", new DataPoints($"{nameof(table)}.Index", table.Index));


            Tracer.Log("Deleting first row");
            table.Rows.First().Delete();

            var headersToRemove = "UL\r\nPower\r\n(dBm)\tUL\r\nS/N\r\n(dB)\tUL\r\nFBER\r\n(%)\tResult\tDL\r\nLoss\r\n(dB)\r\n"
                .Split(new[] { '\t' }, StringSplitOptions.RemoveEmptyEntries)
                .Select(s => FixerHelpers.NormalizeMatchString(s))
                .Reverse()
                .ToList();

            table.Columns
                .Where(col => headersToRemove
                    .Contains(FixerHelpers.NormalizeMatchString(col.Cells[1].Text)))
                .Reverse().ToList().ForEach(col => col.Delete());

            table.AddAndMergeFirstRow("Area Report");
            table.MakeFullPage();

        }

        private void RemoveTestResultReferencesPage2()
        {
            Log.Information("Removing Page 2");
            var tf = new TextFinder(CKDoc, "Threshold Settings");
            if (tf.TryFind(out Word.Range foundRange))
            {
                foundRange.MoveStartUntil(
                    Word.WdBreakType.wdPageBreak,
                    Word.WdConstants.wdBackward);
                foundRange.Start += 2;

                // Find the next section break (instead of using MoveEndUntil)
                Word.Range sectionBreakRange = foundRange.Document.Range(foundRange.End, foundRange.Document.Content.End);
                sectionBreakRange.Find.ClearFormatting();
                sectionBreakRange.Find.Text = "^b"; // Word's special character for section breaks
                sectionBreakRange.Find.Forward = true; // Search forward
                sectionBreakRange.Find.Wrap = Word.WdFindWrap.wdFindStop;

                if (sectionBreakRange.Find.Execute())
                {
                    foundRange.End = sectionBreakRange.End - 1; // Expand the range to the section break
                }
                foundRange.Delete();
                //foundRange.Text = "^b";
            }
        }

        private void FindAndDeleteParagraph(string textToFind)
        {
            var tf = new TextFinder(CKDoc, textToFind);
            if (tf.TryFind(out Word.Range fr))
            {
                fr.Expand(Word.WdUnits.wdParagraph);
                //fr.End = fr.End + 1; // get the paragraph character
                //fr.HighlightColorIndex = Word.WdColorIndex.wdBlue;
                fr.Delete();
            }
        }
        private void RemoveTestResultReferencesPage1()
        {
            Log.Information("...Page 1");
            //Remove:
            //  Result: Fail paragraph
            FindAndDeleteParagraph("(Adjacent Area Rule)");

            //  TestReportSummary columns Result:AreaPointsPassed:CriticalPointsPassed
            //FindAndDeleteParagraph("Test Report Summary");

            string TRSTable_ss = "Channel/ Ch Group\tFreq (MHz)\tTechnology\tBand\tResult\tArea Points\r\npassed (%)\tCritical Points passed (%)\r\n";
            var TRSTable = CKDoc.Tables
                .First(t => t.RowMatches(1, TRSTable_ss));
            TRSTable.Columns[7].Delete();
            TRSTable.Columns[6].Delete();
            TRSTable.Columns[5].Delete();
            TRSTable.MakeFullPage();

            //  TestDetails Cols ResultCalculation:ByAreaPerFloor
            string TDTable_ss = "Test Details";
            var TDTable = CKDoc.Tables
                .First(t => t.RowMatches(1, TDTable_ss));
            FixReportDetailTable(TDTable);
            //TDTable.Columns[4].Delete();
            //TDTable.Columns[3].Delete();
            //TDTable.MakeFullPage();
        }
        private void FixReportDetailTable(CKTable fixer)
        {

            Tracer.Log("Entering", "**", new DataPoints()
                .Add($"{nameof(fixer)}.Index", fixer.Index));

            try
            {
                Tracer.Log("Deleting first row");

                fixer.Rows.First().Delete();


                Tracer.Log("Deleting columns 3 and 4");

                fixer.Columns[4].Delete();
                fixer.Columns[3].Delete();


                Tracer.Log("Returing first row and making fix width");

                fixer.AddAndMergeFirstRow("Test Details");
                fixer.MakeFullPage();
            }
            catch (Exception ex)
            {
                LH.LogThrow(ex);
            }
        }
    }
}
