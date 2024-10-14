//using AKUtilities;
using Microsoft.Office.Interop.Word;

using Word = Microsoft.Office.Interop.Word;

namespace CoverageKiller2

{
    /// <summary>
    /// depreciated
    /// </summary>
    public static class CkDoc
    {
        //        private static string FilePath { get; set; }

        //        #region TableExtensions

        //        public static void DelColumn(this Table table, int colIndex)
        //        {
        //            foreach (Row row in table.Rows)
        //            {
        //                if (row.IsFirst) continue;
        //                //var c = row.Cells[colIndex];
        //                //var r = doc.Range(c.Range.Start, c.Range.End);
        //                //r.Text = "";
        //                row.Cells[colIndex].Delete();
        //            }
        //        }

        //        public static void DelColumn(this Table table, string headingText)
        //        {
        //            var foundColIndex = GetColIndexByText(table, headingText);

        //            DelColumn(table, foundColIndex);
        //        }

        //        public static void DelColumns(this Table table, params string[] columnHeaders)
        //        {
        //            foreach (string cH in columnHeaders)
        //            {
        //                table.DelColumn(cH);
        //            }
        //        }

        //        public static void DelColumnsIf(this Table table, string tableHeader, params string[] columnHeaders)
        //        {
        //            string tableText = table.Range.Text;
        //            if (tableText.Contains(tableHeader))
        //            {
        //                //Remove DL Loss Column
        //                table.DelColumns(columnHeaders);
        //            }
        //        }

        //        public static int GetColIndexByText(this Table table, string headingText)
        //        {
        //            var cells = table.Rows[2].Cells.Cast<Cell>().ToList();
        //            var hT = headingText.TrimWs();

        //            var foundCell = cells.FirstOrDefault(
        //                c => c.Range.Text.TrimWs().Contains(hT));

        //            if (foundCell == null)
        //                throw new Exception("Try again");

        //            return foundCell.ColumnIndex;
        //        }

        //        public static void RenameColumn(this Table table, string oldName, string newName)
        //        {
        //            Cell x = table.Rows[2].Cells[table.GetColIndexByText(oldName)];
        //            x.Range.Text = newName;
        //        }

        //        #endregion TableExtensions

        //        #region Sections

        //        public static void FixDataPageHeadingText(this Document doc, Section section, string newTitle)
        //        {
        //            //color areas to delete
        //            if (section.Range.Paragraphs.First.Range.Text.Contains("Floor: ")
        //                && section.Range.Tables.Count > 0)
        //            {
        //                int delStart = section.Range.Paragraphs[1].Range.End;
        //                int headerEnd = section.Range.Tables[1].Range.Start - 1;
        //                Range rangeToReplace = doc.Range(delStart, headerEnd);

        //                rangeToReplace.Text = newTitle;
        //                rangeToReplace.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
        //            }
        //        }

        //        public static void FixFirstAndLastPages(this Document doc)
        //        {
        //            //clean up 1st page
        //            Word.Section s1 = doc.Sections.First;
        //            //s1.Range.Font.ColorIndex = Word.WdColorIndex.wdViolet;
        //            //s1.Range.Paragraphs.First.Range.Font.ColorIndex = Word.WdColorIndex.wdBrightGreen;

        //            if (s1.Range.Paragraphs.First.Range.Text.Contains("Emergency Responder Radio System"))
        //            {
        //                int selStart = s1.Range.Paragraphs.First.Range.End;
        //                int selEnd = s1.Range.End - 2;
        //                Word.Range r1 = doc.Range(selStart, selEnd);
        //                //r1.Font.ColorIndex = Word.WdColorIndex.wdRed;
        //                r1.Delete();

        //                s1.Range.Paragraphs.First.Range.Text = "PasteFrontMatter\r";

        //                //remove additional info section
        //                doc.Sections.Last.Range.Delete();
        //            }
        //        }

        //        public static bool FixMapDataPageTables(this Table table)
        //        {
        //            var tableText = table.Range.Text.TrimWs();
        //            if (tableText.Contains("Freq (MHz)".TrimWs())
        //                || tableText.Contains("# of Areas".TrimWs())
        //                || tableText.Contains("Reference Point".TrimWs()))
        //            {
        //                table.Delete();
        //                return true;
        //            }
        //            return false;
        //        }

        //        public static List<Section> GetDataPageSections(this Document doc)
        //            => doc.Sections.Cast<Section>()
        //                .Where(s => s.Index > 1)
        //                .ToList();

        //        #endregion Sections

        //        #region HeadersAndFooters

        //        public static void ClearClipboard()
        //        {
        //            //'Clearing the Office Clipboard

        //            //Dim oData   As New DataObject 'object to use the clipboard

        //            //oData.SetText text:= Empty 'Clear
        //            //oData.PutInClipboard 'take in the clipboard to empty it
        //            Clipboard.Clear();
        //        }

        //        //public static void FixHeadersAndFooters(this Document WDA)
        //        //{
        //        //    //Sub PCTEL_ReportHeadersAndFooters()
        //        //    //Dim WDA As Word.Document, WDT As Word.Document
        //        //    Word.Application app = WDA.Application;
        //        //    WDADocument = WDA;

        //        //    FilePath = IndoorReportTemplate.LoadResourceAndCreateTempFile();
        //        //    app.DocumentOpen += App_DocumentOpen;
        //        //    //'get the template
        //        //    //Documents.Open _
        //        //    //FileName:= "C:\Users\akeyte\Desktop\ODIN\Project_Managers\Aaron K\Editing\PCTEL\PCTELReportHeaderFooterTemplate.docx", _
        //        //    //AddToRecentFiles:= False
        //        //    app.Documents.Open(
        //        //        FileName: FilePath,//@"C:\Users\akeyte\source\repos\CKPCTELFix\CKPCTELFix\PCTELReportHeaderFooterTemplate.docx",
        //        //        AddToRecentFiles: false);

        //        //    //now wait for the file to open...
        //        //}
        //        private static Document WDADocument { get; set; }
        //        private static void App_DocumentOpen(Document Doc)
        //        {
        //            Word.Application app = Doc.Application;
        //            //Set WDT = Windows("PCTELReportHeaderFooterTemplate.docx").Document
        //            app.DocumentOpen -= App_DocumentOpen;
        //            if (Doc.FullName != FilePath) return;

        //            string fileName = Path.GetFileName(FilePath);
        //            Word.Document WDT = app.Windows[fileName].Document;
        //            Document WDA = WDADocument;
        //            //'copy headers
        //            //HeaderWholeStory WDT
        //            //Selection.Copy
        //            //WDA.Activate
        //            //Selection.GoTo What:= wdGoToSection, which:= wdGoToAbsolute, Count:= 1
        //            //HeaderWholeStory WDA
        //            //Selection.PasteAndFormat wdFormatOriginalFormatting

        //            //'copy headers
        //            WDT.SelectHeaderWholeStory();
        //            app.Selection.Copy();
        //            WDA.Activate();
        //            app.Selection.GoTo(What: WdGoToItem.wdGoToSection, Which: WdGoToDirection.wdGoToAbsolute, Count: 1);
        //            WDA.SelectHeaderWholeStory();
        //            app.Selection.PasteAndFormat(WdRecoveryType.wdFormatOriginalFormatting);

        //            //'copy footers
        //            //Selection.Copy
        //            //WDA.Activate
        //            //Selection.GoTo What:= wdGoToSection, which:= wdGoToAbsolute, Count:= 1
        //            //FooterWholeStory WDA
        //            //Selection.PasteAndFormat wdFormatOriginalFormatting
        //            WDT.SelectFooterWholeStory();
        //            app.Selection.Copy();
        //            WDA.Activate();
        //            app.Selection.GoTo(What: WdGoToItem.wdGoToSection, Which: WdGoToDirection.wdGoToAbsolute, Count: 1);
        //            WDA.SelectFooterWholeStory();
        //            app.Selection.PasteAndFormat(WdRecoveryType.wdFormatOriginalFormatting);
        //            ClearClipboard();

        //            //'additional fixes
        //            //ClearAdditionalInfo WDA
        //            SetMargins(WDA);
        //            // (not here) ClearAdditionalInfo

        //            //'cleanup
        //            //WDT.Close False
        //            //WDA.Activate
        //            //WDA.ActiveWindow.ActivePane.View.SeekView = wdSeekMainDocument
        //            //WDA.Save
        //            //End Sub

        //            WDT.Close(false);
        //            WDA.Activate();
        //            WDA.ActiveWindow.ActivePane.View.SeekView = WdSeekView.wdSeekMainDocument;
        //            // (not here) WDA.Save
        //        }


        /// <summary>
        /// depreciated
        /// </summary>
        public static void SelectFooterWholeStory(this Document WD)
        {
            Word.Application app = WD.Application;

            //Private Sub FooterWholeStory(WD As Word.Document)
            //WD.Activate
            WD.Activate();

            //If WD.ActiveWindow.View.SplitSpecial<> wdPaneNone Then
            //WD.ActiveWindow.Panes(2).Close
            //End If
            if (WD.ActiveWindow.View.SplitSpecial != WdSpecialPane.wdPaneNone)
            {
                WD.ActiveWindow.Panes[2].Close();
            }

            //If WD.ActiveWindow.ActivePane.View.Type = wdNormalView Or WD.ActiveWindow._
            //ActivePane.View.Type = wdOutlineView Then
            //WD.ActiveWindow.ActivePane.View.Type = wdPrintView
            //End If
            if (WD.ActiveWindow.ActivePane.View.Type == WdViewType.wdNormalView
                || WD.ActiveWindow.ActivePane.View.Type == WdViewType.wdOutlineView)
            {
                WD.ActiveWindow.ActivePane.View.Type = WdViewType.wdPrintView;
            }

            //WD.ActiveWindow.ActivePane.View.SeekView = wdSeekCurrentPageFooter
            //Selection.WholeStory
            //End Sub

            WD.ActiveWindow.ActivePane.View.SeekView = WdSeekView.wdSeekCurrentPageFooter;
            app.Selection.WholeStory();
            //End Sub
        }
        /// <summary>
        /// depreciated
        /// </summary>
        public static void SelectHeaderWholeStory(this Document WD)
        {
            Word.Application app = WD.Application;

            //Private Sub HeaderWholeStory(WD As Word.Document)
            //WD.Activate
            WD.Activate();

            //If WD.ActiveWindow.View.SplitSpecial<> wdPaneNone Then
            //WD.ActiveWindow.Panes(2).Close
            //End If
            if (WD.ActiveWindow.View.SplitSpecial != WdSpecialPane.wdPaneNone)
            {
                WD.ActiveWindow.Panes[2].Close();
            }

            //If WD.ActiveWindow.ActivePane.View.Type = wdNormalView Or WD.ActiveWindow._
            //ActivePane.View.Type = wdOutlineView Then
            //WD.ActiveWindow.ActivePane.View.Type = wdPrintView
            //End If
            if (WD.ActiveWindow.ActivePane.View.Type == WdViewType.wdNormalView
                || WD.ActiveWindow.ActivePane.View.Type == WdViewType.wdOutlineView)
            {
                WD.ActiveWindow.ActivePane.View.Type = WdViewType.wdPrintView;
            }

            //WD.ActiveWindow.ActivePane.View.SeekView = wdSeekCurrentPageHeader
            //Selection.WholeStory
            //End Sub
            WD.ActiveWindow.ActivePane.View.SeekView = WdSeekView.wdSeekCurrentPageHeader;
            app.Selection.WholeStory();
        }

        //        public static void SetMargins(this Document WD)
        //        {
        //            Word.Application app = WD.Application;

        //            //Private Sub SetMargins(WD As Word.Document)
        //            //WD.Activate
        //            //WD.ActiveWindow.ActivePane.View.SeekView = wdSeekMainDocument

        //            WD.Activate();
        //            WD.ActiveWindow.ActivePane.View.SeekView = WdSeekView.wdSeekMainDocument;

        //            //Selection.WholeStory
        //            app.Selection.WholeStory();

        //            //With WD.Styles(wdStyleNormal).Font
        //            var font = WD.Styles[WdBuiltinStyle.wdStyleNormal].Font;

        //            //If.NameFarEast = .NameAscii Then
        //            //.NameAscii = ""
        //            //End If
        //            //.NameFarEast = ""
        //            //End With
        //            if (font.NameFarEast == font.NameAscii)
        //                font.NameAscii = "";
        //            font.NameFarEast = "";

        //            //With WD.PageSetup
        //            var pS = WD.PageSetup;

        //            pS.LineNumbering.Active = 0;
        //            pS.Orientation = WdOrientation.wdOrientPortrait;
        //            pS.TopMargin = app.InchesToPoints(1.25f);
        //            pS.BottomMargin = app.InchesToPoints(1f);
        //            pS.LeftMargin = app.InchesToPoints(0.5f);
        //            pS.RightMargin = app.InchesToPoints(0.5f);
        //            pS.Gutter = app.InchesToPoints(0f);
        //            pS.HeaderDistance = app.InchesToPoints(0.5f);
        //            pS.FooterDistance = app.InchesToPoints(0.5f);
        //            pS.PageWidth = app.InchesToPoints(8.5f);
        //            pS.PageHeight = app.InchesToPoints(11f);
        //            pS.FirstPageTray = WdPaperTray.wdPrinterDefaultBin;
        //            pS.OtherPagesTray = WdPaperTray.wdPrinterDefaultBin;
        //            pS.OddAndEvenPagesHeaderFooter = 0;
        //            pS.DifferentFirstPageHeaderFooter = 0;
        //            pS.VerticalAlignment = WdVerticalAlignment.wdAlignVerticalTop;
        //            pS.SuppressEndnotes = 0;
        //            pS.MirrorMargins = 0;
        //            pS.TwoPagesOnOne = false;
        //            pS.BookFoldPrinting = false;
        //            pS.BookFoldRevPrinting = false;
        //            pS.BookFoldPrintingSheets = 1;
        //            pS.GutterPos = WdGutterStyle.wdGutterPosLeft;

        //            //End With
        //            //End Sub
        //        }

        //        #endregion HeadersAndFooters
    }
}