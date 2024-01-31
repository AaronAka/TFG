using Microsoft.Office.Interop.Word;
using System.Collections.Generic;
using Range = Microsoft.Office.Interop.Word.Range;

namespace TFG.Model.ParagraphFormatStrategy
{
    public class SourcesParagraphFormat : IParagraphFormat
    {
        public void formatParagraph(List<string> textWithTags, Document doc)
        {
            object oCollapseEnd = WdCollapseDirection.wdCollapseEnd;
            doc.Paragraphs.Add();
            Range objRange = doc.Content;

            for (int i = 0; i < textWithTags.Count; i++)
            {
                objRange.Collapse(ref oCollapseEnd);
                
                if (i == 0 || i == 8 || i == 10 || i == 12 || i == 15)
                {
                    objRange.Text = textWithTags[i];
                    objRange.Bold = 0;
                    objRange.Font.Name = "Arial";
                    objRange.Font.Size = 12;
                    objRange.Font.Color = WdColor.wdColorGreen;
                    objRange.Paragraphs.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
                }
                else if (i == 1 || i == 3)
                {
                    objRange.Text = textWithTags[i];
                    objRange.Bold = 0;
                    objRange.Font.Name = "Arial";
                    objRange.Font.Size = 12;
                    objRange.Font.Color = WdColor.wdColorRed;
                    objRange.Paragraphs.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
                }
                else if (i == 2 || i == 5 || i == 7 || i == 9 || i == 11 || i == 13 || i == 14)
                {
                    objRange.Text = textWithTags[i];
                    objRange.Bold = 0;
                    objRange.Font.Name = "Times New Roman";
                    objRange.Font.Size = 12;
                    objRange.Font.Color = WdColor.wdColorBlack;
                    objRange.Paragraphs.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
                }
                else if (i == 4 || i == 6)
                {
                    objRange.Text = textWithTags[i];
                    objRange.Bold = 0;
                    objRange.Font.Name = "Arial";
                    objRange.Font.Size = 12;
                    objRange.Font.Color = WdColor.wdColorBlue;
                    objRange.Paragraphs.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
                }
            }
        }
    }
}
