using Microsoft.Office.Interop.Word;
using System.Collections.Generic;
using Range = Microsoft.Office.Interop.Word.Range;

namespace TFG.Model.ParagraphFormatStrategy
{
    public class IntroParagraphFormat : IParagraphFormat
    {
        public void formatParagraph(List<string> textWithTags, Document doc)
        {
            object oCollapseEnd = WdCollapseDirection.wdCollapseEnd;
            doc.Paragraphs.Add();
            Range objRange = doc.Content;

            for (int i = 0; i < textWithTags.Count; i++)
            {
                objRange.Collapse(ref oCollapseEnd);

                objRange.Font.Size = 11;
                objRange.Paragraphs.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;

                if (i == 0)
                {
                    objRange.Text = textWithTags[i];
                    objRange.Font.Name = "Arial";
                    objRange.Font.Color = WdColor.wdColorGreen;
                    
                }
                else if (i == 1 || i == 3)
                {
                    objRange.Text = textWithTags[i];
                    objRange.Font.Name = "Arial";
                    objRange.Font.Color = WdColor.wdColorBlue;
                }
                else
                {
                    objRange.Text = textWithTags[i];
                    objRange.Font.Name = "Times New Roman";
                    objRange.Font.Color = WdColor.wdColorBlack;
                }
            }
        }
    }
}
