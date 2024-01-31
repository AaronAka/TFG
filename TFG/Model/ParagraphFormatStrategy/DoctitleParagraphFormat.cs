using Microsoft.Office.Interop.Word;
using System.Collections.Generic;
using Range = Microsoft.Office.Interop.Word.Range;

namespace TFG.Model.ParagraphFormatStrategy
{
    public class DoctitleParagraphFormat : IParagraphFormat
    {
        public void formatParagraph(List<string> textWithTags, Document doc)
        {
            object oCollapseEnd = WdCollapseDirection.wdCollapseEnd;
            doc.Paragraphs.Add();
            Range objRange = doc.Content;
            var tags = "[doctitle] [/doctitle]";

            for (int i = 0; i < textWithTags.Count; i++)
            {
                objRange.Collapse(ref oCollapseEnd);

                objRange.Text = textWithTags[i];
                objRange.Font.Size = 16;
                objRange.Paragraphs.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;

                if (tags.Contains(textWithTags[i]))
                {
                    objRange.Font.Name = "Arial";
                    objRange.Font.Color = WdColor.wdColorPlum;
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
