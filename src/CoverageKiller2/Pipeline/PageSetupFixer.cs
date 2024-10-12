//using Microsoft.Office.Interop.Word;
using Serilog;
using System.Collections.Generic;
using Word = Microsoft.Office.Interop.Word;
namespace CoverageKiller2.Pipeline
{
    internal class PageSetupFixer : CKWordPipelineProcess
    {
        private IndoorReportTemplate template;

        public PageSetupFixer(IndoorReportTemplate template)
        {
            this.template = template;
        }
        public override void Process()
        {
            Log.Information("Fixing PageSetup");
            CopyPageSetup(template.WordDoc, CKDoc.WordDoc);
            CopyCompleteStyles(template.WordDoc, CKDoc.WordDoc);

            WordSelector.MainDocument(CKDoc);

            //var font = CKDoc.WordDoc.Styles[Word.WdBuiltinStyle.wdStyleNormal].Font;

            //if (font.NameFarEast == font.NameAscii)
            //    font.NameAscii = "";
            //font.NameFarEast = "";
            ////CKDoc.WordDoc.PageSetup = template.WordDoc.PageSetup.;

            //// in future would be good to POCO this to an XML config.
            //var pS = CKDoc.WordDoc.PageSetup;
            //var app = CKDoc.WordDoc.Application;

            //pS.LineNumbering.Active = 0;
            //pS.Orientation = Word.WdOrientation.wdOrientPortrait;
            //pS.TopMargin = app.InchesToPoints(1.25f);
            //pS.BottomMargin = app.InchesToPoints(1f);
            //pS.LeftMargin = app.InchesToPoints(0.5f);
            //pS.RightMargin = app.InchesToPoints(0.5f);
            //pS.Gutter = app.InchesToPoints(0f);
            //pS.HeaderDistance = app.InchesToPoints(0.5f);
            //pS.FooterDistance = app.InchesToPoints(0.5f);
            //pS.PageWidth = app.InchesToPoints(8.5f);
            //pS.PageHeight = app.InchesToPoints(11f);
            //pS.FirstPageTray = Word.WdPaperTray.wdPrinterDefaultBin;
            //pS.OtherPagesTray = Word.WdPaperTray.wdPrinterDefaultBin;
            //pS.OddAndEvenPagesHeaderFooter = 0;
            //pS.DifferentFirstPageHeaderFooter = 0;
            //pS.VerticalAlignment = Word.WdVerticalAlignment.wdAlignVerticalTop;
            //pS.SuppressEndnotes = 0;
            //pS.MirrorMargins = 0;
            //pS.TwoPagesOnOne = false;
            //pS.BookFoldPrinting = false;
            //pS.BookFoldRevPrinting = false;
            //pS.BookFoldPrintingSheets = 1;
            //pS.GutterPos = Word.WdGutterStyle.wdGutterPosLeft;

        }



        public void CopyPageSetup(Word.Document sourceDoc, Word.Document targetDoc)
        {
            Log.Debug("Copying Page Setup.");
            // Access the PageSetup of the source document
            var sourcePageSetup = sourceDoc.PageSetup;

            // Access the PageSetup of the target document
            var targetPageSetup = targetDoc.PageSetup;

            // Copy properties from the source to the target
            targetPageSetup.Orientation = sourcePageSetup.Orientation;
            targetPageSetup.TopMargin = sourcePageSetup.TopMargin;
            targetPageSetup.BottomMargin = sourcePageSetup.BottomMargin;
            targetPageSetup.LeftMargin = sourcePageSetup.LeftMargin;
            targetPageSetup.RightMargin = sourcePageSetup.RightMargin;
            //targetPageSetup.HeaderDistance = sourcePageSetup.HeaderDistance;
            //targetPageSetup.FooterDistance = sourcePageSetup.FooterDistance;
            targetPageSetup.PageHeight = sourcePageSetup.PageHeight;
            targetPageSetup.PageWidth = sourcePageSetup.PageWidth;
            targetPageSetup.Gutter = sourcePageSetup.Gutter;
            targetPageSetup.SuppressEndnotes = sourcePageSetup.SuppressEndnotes;
            //targetPageSetup.EvenAndOddHeaders = sourcePageSetup.EvenAndOddHeaders;
            targetPageSetup.DifferentFirstPageHeaderFooter = sourcePageSetup.DifferentFirstPageHeaderFooter;

            // Copy line numbering settings
            var sourceLineNumbering = sourcePageSetup.LineNumbering;
            targetPageSetup.LineNumbering.Active = sourceLineNumbering.Active;
            targetPageSetup.LineNumbering.CountBy = sourceLineNumbering.CountBy;
            targetPageSetup.LineNumbering.StartingNumber = sourceLineNumbering.StartingNumber;
            //targetPageSetup.LineNumbering.RestartNumberingAtStart = sourceLineNumbering.RestartNumberingAtStart;

            // Copy other relevant properties
            targetPageSetup.PaperSize = sourcePageSetup.PaperSize;
            //targetPageSetup.AllowPageBreaks = sourcePageSetup.AllowPageBreaks;
            targetPageSetup.LineNumbering.Active = sourceLineNumbering.Active;
            targetPageSetup.VerticalAlignment = sourcePageSetup.VerticalAlignment;

            // Add any additional properties you wish to copy

            Log.Debug("Done copying Page Setup.");
        }




        public void CopyCompleteStyles(Word.Document sourceDoc, Word.Document targetDoc)
        {
            Log.Information("Copying Styles...");
            int copiedStylesCount = 0;
            var styleNames = new List<string>();

            foreach (Word.Style sourceStyle in sourceDoc.Styles)
            {

                Word.Style targetStyle = targetDoc.Styles[sourceStyle.NameLocal];
                styleNames.Add(sourceStyle.NameLocal);
                Log.Debug("Copying {SourceStyle}...", sourceStyle.NameLocal);


                //debug
                //if (sourceStyle.NameLocal.Contains("Block"))
                //{
                //    Log.Debug("/tSkipped {StyleName}", sourceStyle.NameLocal);
                //    continue;
                //}


                if (targetStyle != null)
                {

                    Log.Debug("\tFont...");
                    // Copy font settings
                    targetStyle.Font.Name = sourceStyle.Font.Name;
                    targetStyle.Font.Size = sourceStyle.Font.Size;
                    targetStyle.Font.Bold = sourceStyle.Font.Bold;
                    targetStyle.Font.Italic = sourceStyle.Font.Italic;
                    targetStyle.Font.Underline = sourceStyle.Font.Underline;
                    targetStyle.Font.Color = sourceStyle.Font.Color;
                    targetStyle.Font.StrikeThrough = sourceStyle.Font.StrikeThrough;
                    targetStyle.Font.Superscript = sourceStyle.Font.Superscript;
                    targetStyle.Font.Subscript = sourceStyle.Font.Subscript;
                    targetStyle.Font.SmallCaps = sourceStyle.Font.SmallCaps;
                    targetStyle.Font.AllCaps = sourceStyle.Font.AllCaps;
                    targetStyle.Font.Hidden = sourceStyle.Font.Hidden;

                    // Check if the style is a paragraph style
                    if (sourceStyle.Type == Word.WdStyleType.wdStyleTypeParagraph)
                    {
                        Log.Debug("\tParagraph Format... 0");
                        // Copy paragraph format properties
                        targetStyle.ParagraphFormat.Alignment = sourceStyle.ParagraphFormat.Alignment;
                        targetStyle.ParagraphFormat.LineSpacing = sourceStyle.ParagraphFormat.LineSpacing;
                        targetStyle.ParagraphFormat.SpaceBefore = sourceStyle.ParagraphFormat.SpaceBefore;
                        Log.Debug("\t\t...3");
                        targetStyle.ParagraphFormat.SpaceAfter = sourceStyle.ParagraphFormat.SpaceAfter;
                        targetStyle.ParagraphFormat.LeftIndent = sourceStyle.ParagraphFormat.LeftIndent;
                        targetStyle.ParagraphFormat.RightIndent = sourceStyle.ParagraphFormat.RightIndent;
                        Log.Debug("\t\t...6");
                        targetStyle.ParagraphFormat.FirstLineIndent = sourceStyle.ParagraphFormat.FirstLineIndent;
                        targetStyle.ParagraphFormat.KeepTogether = sourceStyle.ParagraphFormat.KeepTogether;
                        targetStyle.ParagraphFormat.KeepWithNext = sourceStyle.ParagraphFormat.KeepWithNext;
                        Log.Debug("\t\t...9");
                        targetStyle.ParagraphFormat.PageBreakBefore = sourceStyle.ParagraphFormat.PageBreakBefore;
                        targetStyle.ParagraphFormat.WidowControl = sourceStyle.ParagraphFormat.WidowControl;
                        Log.Debug("\t\t...11");
                        //targetStyle.ParagraphFormat.ContextualSpacing = sourceStyle.ParagraphFormat.ContextualSpacing;

                        // Copy borders
                        //targetStyle.ParagraphFormat.Borders.Enable = sourceStyle.ParagraphFormat.Borders.Enable;
                        Log.Debug("\t\t...12(skipped)");
                        //targetStyle.ParagraphFormat.Borders.OutsideLineStyle = sourceStyle.ParagraphFormat.Borders.OutsideLineStyle;
                        Log.Debug("\t\t...13(skipped)");
                        //targetStyle.ParagraphFormat.Borders.OutsideColor = sourceStyle.ParagraphFormat.Borders.OutsideColor;

                        // Copy shading
                        targetStyle.ParagraphFormat.Shading.BackgroundPatternColor = sourceStyle.ParagraphFormat.Shading.BackgroundPatternColor;
                        Log.Debug("\t\t...14 (done)");

                        // Copy numbering
                        //targetStyle.ParagraphFormat.NumberingFormat = sourceStyle.ParagraphFormat.NumberingFormat;

                        // Copy tab stops
                        Log.Debug("\tTab stops...");

                        foreach (Word.TabStop tabStop in sourceStyle.ParagraphFormat.TabStops)
                        {
                            targetStyle.ParagraphFormat.TabStops.Add(tabStop.Position, tabStop.Alignment);
                        }
                    }
                    else if (sourceStyle.Type == Word.WdStyleType.wdStyleTypeCharacter)
                    {
                        Log.Debug("\tCharacter Type (Not Copied)");
                        // Handle character styles (if necessary)
                        // For now, copying font settings is sufficient
                    }
                }
                copiedStylesCount++;

            }
            Log.Information("Processed {Count} styles: \n{StyleNames}", copiedStylesCount, styleNames);
        }
    }
}