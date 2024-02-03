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

                objRange.Text = textWithTags[i];
                objRange.Bold = 0;
                objRange.Font.Size = 12;
                objRange.Paragraphs.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;

                if (textWithTags[i].Contains("normaff") || textWithTags[i].Contains("city") || textWithTags[i].Contains("country"))
                {
                    objRange.Font.Name = "Arial";
                    objRange.Font.Color = WdColor.wdColorGreen;
                }
                else if (textWithTags[i].Contains("label") || textWithTags[i].Contains("sup"))
                {
                    objRange.Font.Name = "Arial";
                    objRange.Font.Color = WdColor.wdColorRed;
                }
                else if (textWithTags[i].Contains("orgdiv"))
                {
                    objRange.Font.Name = "Arial";
                    objRange.Font.Color = WdColor.wdColorBlue;
                }
                else
                {
                    objRange.Font.Name = "Times New Roman";
                    objRange.Font.Color = WdColor.wdColorBlack;
                }
            }
        }
    }
}
