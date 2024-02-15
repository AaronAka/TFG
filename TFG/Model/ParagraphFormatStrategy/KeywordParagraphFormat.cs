using Microsoft.Office.Interop.Word;
using System.Collections.Generic;
using Range = Microsoft.Office.Interop.Word.Range;

namespace TFG.Model.ParagraphFormatStrategy
{
    public class KeywordParagraphFormat : IParagraphFormat
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
                    objRange.Font.Color = WdColor.wdColorBrown;
                }
                else if (i == 1 || i == 3)
                {
                    objRange.Font.Name = "Arial";
                    objRange.Font.Color = WdColor.wdColorBlue;
                }
                else if (textWithTags[i] == "[kwd]" || textWithTags[i].Contains("[/kwd]"))
                {
                    if (textWithTags[i] == "[kwd]" && !textWithTags[i - 1].Contains("[/kwd]"))
                    {
                        objRange.Text = "\t" + textWithTags[i];
                    }
                    else
                    {
                        objRange.Text = textWithTags[i];
                    }

                    objRange.Font.Name = "Arial";
                    objRange.Font.Color = WdColor.wdColorGreen;
                }
                else
                {
                    objRange.Text = textWithTags[i].Trim();
                    objRange.Font.Name = "Times New Roman";
                    objRange.Font.Color = WdColor.wdColorBlack;
                }
            }
        }
    }
}