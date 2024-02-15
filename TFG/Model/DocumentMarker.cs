using Microsoft.Office.Interop.Word;
using System;
using System.Collections.Generic;
using System.IO;
using System.Windows;
using TFG.Model.ParagraphFormatStrategy;
using Application = Microsoft.Office.Interop.Word.Application;

namespace TFG.Model
{
    public class DocumentMarker
    {
        private IParagraphFormat strategy;
        private int tableIndex = 1;
        private int refIndex = 1;
        private bool inBody = false;
        private bool reachedBibliography = false;
        private bool colourString = true;

        public DocumentMarker()
        {
            strategy = new DoctitleParagraphFormat();
        }

        public void MarkDocument(string documentRoute)
        {
            if (documentRoute != null)
            {
                Object miss = Type.Missing;
                object readOnly = true;
                object isVisible = false;

                Application application = new Application();
                Document document = application.Documents.Add(documentRoute);
                Document newDocument = application.Documents.Add();

                List<string> extractedWords = new List<string>();
                foreach (Paragraph paragraph in document.Paragraphs)
                {
                    try
                    {
                        extractedWords.Clear();
                        if (ParagraphTypeIdentifier.EmptyParagraph(paragraph))
                        {
                            continue;
                        }

                        if (reachedBibliography)
                        {
                            PrepareAndInsertBibliographyTags(extractedWords, paragraph);
                        }
                        else
                        {
                            ParagraphType paragraphType = ParagraphTypeIdentifier.GetParagraphType(inBody, paragraph);
                            InsertTags(paragraphType, extractedWords, paragraph);
                        }

                        if (extractedWords.Count > 0)
                        {
                            strategy.formatParagraph(extractedWords, newDocument);
                        }
                    }
                    catch
                    {

                    }
                }

                SaveMarkedDocument(documentRoute, ref miss, ref application, ref document, newDocument);
            }
        }

        private static void SaveMarkedDocument(string documentRoute, ref object miss, ref Application application, ref Document document, Document newDocument)
        {
            try
            {
                object fileName = Path.GetDirectoryName(documentRoute) + "\\" + Path.GetFileNameWithoutExtension(documentRoute) + "-marked.docx";
                //object fileName = new FileInfo(documentRoute).Directory.FullName + System.IO.Path.GetFileName(documentRoute) + "-marked.docx";

                newDocument.PageSetup.TopMargin = application.InchesToPoints((float)0.5);
                newDocument.PageSetup.BottomMargin = application.InchesToPoints((float)0.5);
                newDocument.PageSetup.LeftMargin = application.InchesToPoints((float)0.5);
                newDocument.PageSetup.RightMargin = application.InchesToPoints((float)0.5);
                application.DisplayAlerts = WdAlertLevel.wdAlertsNone;
                newDocument.SaveAs2(ref fileName);
                document.Close(ref miss, ref miss, ref miss);
                document = null;
                application.Quit(ref miss, ref miss, ref miss);
                application = null;

                MessageBox.Show("The document has been marked successfully");
            }
            catch (Exception e)
            {
                if (e is IOException)
                {
                    MessageBox.Show("The target file can't be accessed, please make sure that it isn't opened by another processs.");
                }
            }
        }

        private void PrepareAndInsertBibliographyTags(List<string> extractedWords, Paragraph paragraph)
        {
            extractedWords.Add(string.Format(MarkingConstants.BIBLIOGRAPHY_REF_OPEN, refIndex));
            var text = paragraph.Range.Text.Replace("\r", "").Replace("\u000e", "");
            var splitText = text.Split(',');

            MarkDocumentFunctions.AddBibliographyTags(extractedWords, splitText);

            //taggedString += "[/ref]";
            extractedWords.Add(MarkingConstants.BIBLIOGRAPHY_REF_CLOSE);

            refIndex++;
            SetStrategy(new BibliographyParagraphFormat());
        }

        private void InsertTags(ParagraphType paragraphType, List<string> extractedWords, Paragraph paragraph)
        {
            switch (paragraphType)
            {
                case ParagraphType.Empty:
                    break;

                case ParagraphType.Doctitle:
                    MarkDocumentFunctions.AddDoctitleTags(extractedWords, paragraph);
                    SetStrategy(new DoctitleParagraphFormat());
                    break;

                case ParagraphType.Authors:
                    MarkDocumentFunctions.MarkAuthorsInterop(paragraph.Range.Text, extractedWords);
                    SetStrategy(new AuthorsParagraphFormat());
                    break;

                case ParagraphType.Sectitle:
                    MarkDocumentFunctions.AddSecTypeTags(extractedWords, paragraph, inBody);
                    if (!inBody)
                    {
                        inBody = extractedWords[0].Contains("xmlbody");
                    }
                    colourString = paragraph.Range.Text != "MÉTODO\r";
                    SetStrategy(new IntroParagraphFormat());
                    break;

                case ParagraphType.Subsection:
                    SubSecPrep(inBody, extractedWords, paragraph);
                    reachedBibliography = inBody && extractedWords[0].Contains("BIBLIOGRAFÍA");
                    break;

                case ParagraphType.Regular:
                    RegularPrep(colourString, extractedWords, paragraph.Range.Text.Replace("\r", "").Replace("\u000e", ""));
                    break;

                case ParagraphType.Table:
                    MarkDocumentFunctions.AddTableTags(tableIndex, extractedWords, paragraph);
                    tableIndex++;
                    SetStrategy(new TableParagraphFormat());
                    break;

                case ParagraphType.Keywords:
                    MarkDocumentFunctions.AddKeywordTags(extractedWords, paragraph);
                    SetStrategy(new KeywordParagraphFormat());
                    break;

                case ParagraphType.Sources:
                    MarkDocumentFunctions.AddSourceTags(extractedWords, paragraph);
                    SetStrategy(new SourcesParagraphFormat());
                    break;
            }
        }

        private void SubSecPrep(bool inBody, List<string> extractedWords, Paragraph paragraph)
        {
            if (!inBody)
            {
                extractedWords.Add(string.Format(MarkingConstants.SECTITLE, paragraph.Range.Text.Replace("\r", "")));

                SetStrategy(new SectitleParagraphFormat());
            }
            else
            {
                extractedWords.Add("[subsec]" + string.Format(MarkingConstants.SECTITLE, paragraph.Range.Text.Replace("\r", "")));

                SetStrategy(new RegularParagraphFormat());
            }
        }

        private void RegularPrep(bool colourString, List<string> extractedWords, string rawText)
        {
            if (!colourString)
            {
                extractedWords.Add(string.Format(MarkingConstants.REGULAR_PARAGRAPH, rawText));

                SetStrategy(new RegularParagraphFormat());
            }
            else
            {
                extractedWords.Add(MarkingConstants.REGULAR_PARAGRAPH_OPEN);
                extractedWords.Add(rawText);
                extractedWords.Add(MarkingConstants.REGULAR_PARAGRAPH_CLOSE);

                SetStrategy(new RegularParagraphColoured());
            }
        }

        private void SetStrategy(IParagraphFormat strategy)
        {
            this.strategy = strategy;
        }
    }
}