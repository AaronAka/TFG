using Microsoft.Office.Interop.Word;
using System.Collections.Generic;
using Range = Microsoft.Office.Interop.Word.Range;

namespace TFG.Model.ParagraphFormatStrategy
{
    public class RegularParagraphColoured : IParagraphFormat
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
                objRange.Font.Size = 11;
                objRange.Paragraphs.Alignment = WdParagraphAlignment.wdAlignParagraphJustify;

                if (i == 0 || i == textWithTags.Count - 1)
                {
                    objRange.Font.Name = "Arial";
                    objRange.Font.Color = WdColor.wdColorRed;
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