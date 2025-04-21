using CoverageKiller2.DOM;
using CoverageKiller2.DOM.Tables;
using CoverageKiller2.Logging;
using Serilog;
using System;
using System.Diagnostics;
using System.Linq;

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
        private bool _caughtException = false;

        public SEA2025Fixer()
        {
            Tracer.Enabled = true;
        }

        public Tracer Tracer { get; } = new Tracer(typeof(SEA2025Fixer));
        public override void Process()
        {
            this.Ping();

            try
            {

                Log.Information("**** Fixing for SEA2025");

                //remove test result references
                RemoveTestResultRefernces();
                //Add test radio information

            }
            catch (CKDebugException ex)
            {
                LH.Error(ex, rethrow: false);
            }
            catch (Exception ex)
            {
                LH.Error(ex, rethrow: true);
            }
            this.Pong();

        }

        private void RemoveTestResultRefernces()
        {
            this.Ping();

            Log.Information("Removing Test Result References");
            //Remove Page 1 result references
            RemoveTestResultReferencesPage1();
            //Remove Page 2
            RemoveTestResultReferencesPage2();
            //Remove Floor Section result references
            RemoveTestResultReferencesFloorSections();
            //Remove the Informaton
            RemoveMoreInfoSection();
            this.Pong();

        }

        private void RemoveMoreInfoSection()

        {
            this.Ping();

            var x = CKDoc.Range().TryFindNext("Additional Info");
            x.Sections[1].Delete();

            this.Pong();

            //var tf = new TextFinder(CKDoc, );
            //if (tf.TryFind(out Word.Range foundRange))
            //{
            //    foundRange = FixerHelpers.FindSectionBreak(foundRange, false);
            //    foundRange.End = foundRange.Document.Content.End;
            //    foundRange.HighlightColorIndex = Word.WdColorIndex.wdBlue;
            //    foundRange.Delete();

            //}
        }
        private void RemoveTestResultReferencesFloorSections()
        {
            this.Ping();

            Log.Information("Deleting Floor Section pass/fail subtitle.");
            var x = CKDoc.Content.TryFindNext("Result: *", matchWildcards: true);
            x.Paragraphs[1].Delete();
            this.Pong();


            //var tf1 = new TextFinder(CKDoc, );
            //while (tf1.TryFind(out var fr1, true))
            //{
            //    fr1.Expand(Word.WdUnits.wdParagraph);
            //    fr1.Delete();
            //}

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
            this.Ping();

            var headersToRemove = "Area Points passed (%)\tCritical Points passed (%)"
                .Split(new[] { '\t' }, StringSplitOptions.RemoveEmptyEntries)
                .Select(s => FixerHelpers.NormalizeMatchString(s))
                .Reverse()
                .ToList();

            table.Columns
                .Where(col => headersToRemove
                    .Contains(FixerHelpers.NormalizeMatchString(col[1].Text)))
                .Reverse().ToList().ForEach(col => col.Delete());

            table.MakeFullPage();
            this.Pong();

        }

        private void FixFloorSectionCriticalPointReportTable(CKTable table)
        {
            this.Ping();



            var headersToRemove = "UL\r\nPower\r\n(dBm)\tUL\r\nS/N\r\n(dB)\tUL\r\nFBER\r\n(%)\tResult\tDL\r\nLoss\r\n(dB)\r\n"
                .Split(new[] { '\t' }, StringSplitOptions.RemoveEmptyEntries)
                .Select(s => FixerHelpers.NormalizeMatchString(s))
                .Reverse()
                .ToList();

            table.Columns
                .Where(col => headersToRemove
                    .Contains(FixerHelpers.NormalizeMatchString(col[1].Text)))
                .Reverse().ToList().ForEach(col => col.Delete());

            table.MakeFullPage();
            this.Pong();

        }
        private void FixFloorSectionAreaReportTable(CKTable table)
        {
            this.Ping();

            var headersToRemove = "UL\r\nPower\r\n(dBm)\tUL\r\nS/N\r\n(dB)\tUL\r\nFBER\r\n(%)\tResult\tDL\r\nLoss\r\n(dB)\r\n"
                .Split(new[] { '\t' }, StringSplitOptions.RemoveEmptyEntries)
                .Select(s => FixerHelpers.NormalizeMatchString(s))
                .Reverse()
                .ToList();

            table.Columns
                .Where(col => headersToRemove
                    .Contains(FixerHelpers.NormalizeMatchString(col[1].Text)))
                .Reverse().ToList().ForEach(col => col.Delete());

            table.MakeFullPage();
            this.Pong();

        }

        private void RemoveTestResultReferencesPage2()
        {
            this.Ping();

            Log.Information("Removing Page 2");
            LH.Checkpoint("Section 2 delete", GetType());

            //var x = 
            CKDoc.Sections[2].Delete();
            //x.Shading.BackgroundPatternColor = Word.WdColor.wdColorBlue;
            //var y = CKDoc.COMObject.Range().Sections[3].Range;
            //y.Shading.BackgroundPatternColor = Word.WdColor.wdColorAqua;
            //var z = CKDoc.COMObject.Range().Sections[4].Range;
            //z.Shading.BackgroundPatternColor = Word.WdColor.wdColorSeaGreen;
            //x.Text = "";

            //var tf = new TextFinder(CKDoc, "Threshold Settings");

            //if (tf.TryFind(out Word.Range foundRange))
            //{


            //    foundRange.MoveStartUntil(
            //        Word.WdBreakType.wdPageBreak,
            //        Word.WdConstants.wdBackward);
            //    foundRange.Start += 2;

            //    // Find the next section break (instead of using MoveEndUntil)
            //    Word.Range sectionBreakRange = foundRange.Document.Range(foundRange.End, foundRange.Document.Content.End);
            //    sectionBreakRange.Find.ClearFormatting();
            //    sectionBreakRange.Find.Text = "^b"; // Word's special character for section breaks
            //    sectionBreakRange.Find.Forward = true; // Search forward
            //    sectionBreakRange.Find.Wrap = Word.WdFindWrap.wdFindStop;

            //    if (sectionBreakRange.Find.Execute())
            //    {
            //        foundRange.End = sectionBreakRange.End - 1; // Expand the range to the section break
            //    }
            //    //foundRange.Shading.BackgroundPatternColor = Word.WdColor.wdColorBlue;
            //    //foundRange.Delete();
            //    //foundRange.Text = "^b";
            //}
            this.Pong();

        }

        //private void FindAndDeleteParagraph(string textToFind)
        //{



        //    var tf = new TextFinder(CKDoc, textToFind);
        //    if (tf.TryFind(out Word.Range fr))
        //    {
        //        fr.Expand(Word.WdUnits.wdParagraph);
        //        //fr.End = fr.End + 1; // get the paragraph character
        //        //fr.HighlightColorIndex = Word.WdColorIndex.wdBlue;
        //        fr.Delete();
        //    }
        //}
        private void RemoveTestResultReferencesPage1()
        {
            this.Ping();
            try
            {

                Log.Information("...Page 1");

                LH.Checkpoint("(Adjacent Area Rule)", GetType());

                var x = CKDoc.Content.TryFindNext("(Adjacent Area Rule)")
                    ?? CKDoc.Sections[1].TryFindNext("Result: Passed");

                x.Paragraphs[1].Delete();


                //FindAndDeleteParagraph("(Adjacent Area Rule)");

                //  TestReportSummary columns Result:AreaPointsPassed:CriticalPointsPassed
                //FindAndDeleteParagraph("Test Report Summary");

                string TRSTable_ss = "Channel/ Ch Group\tFreq (MHz)\tTechnology\tBand\tResult\tArea Points\r\npassed (%)\tCritical Points passed (%)\r\n";

                LH.Checkpoint(TRSTable_ss, GetType());

                var TRSTable = CKDoc.Tables
                    .First(t =>
                    CKTextHelper.ScrunchEquals(
                        string.Join(string.Empty, t.Rows[2].Select(c => c.Text)),
                        TRSTable_ss)); //throw if null ok

                if (Debugger.IsAttached) Debugger.Break();

                TRSTable.Columns[7].Delete();
                TRSTable.Columns[6].Delete();
                TRSTable.Columns[5].Delete();
                TRSTable.MakeFullPage();


                //  TestDetails Cols ResultCalculation:ByAreaPerFloor
                string TDTable_ss = "Test Details";
                //var TDTable = CKDoc.Tables
                //    .First(t => t.RowMatches(1, TDTable_ss));

                LH.Checkpoint(TDTable_ss, GetType());

                var TDTable = CKDoc.Tables
                    .First(t =>
                    CKTextHelper.ScrunchEquals(
                        string.Join(string.Empty, t.Rows[2].Select(c => c.Text)),
                        TDTable_ss)); //throw if null ok



                FixReportDetailTable(TDTable);
                //TDTable.Columns[4].Delete();
                //TDTable.Columns[3].Delete();
                //TDTable.MakeFullPage();
            }
            catch (Exception ex)
            {
                _caughtException = true;
                LH.Error(ex);
            }
            finally
            {
                //if (Debugger.IsAttached) Debugger.Break();
                if (_caughtException) CKOffice_Word.Instance.Crash(GetType());
            }
            this.Pong();

        }
        private void FixReportDetailTable(CKTable fixer)
        {
            this.Ping();


            Tracer.Log("Entering", "**", new DataPoints()
                .Add($"{nameof(fixer)}.Index", CKDoc.Tables.IndexOf(fixer)));

            try
            {
                Tracer.Log("Deleting columns 3 and 4");

                fixer.Columns[4].Delete();
                fixer.Columns[3].Delete();

                fixer.MakeFullPage();
            }
            catch (Exception ex)
            {
                LH.LogThrow(ex);
            }
            this.Pong();
        }
    }
}
