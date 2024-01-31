using Microsoft.Office.Interop.Word;
using System.Collections.Generic;
using Range = Microsoft.Office.Interop.Word.Range;

namespace TFG.Model.ParagraphFormatStrategy
{
    public class SectitleParagraphFormat : IParagraphFormat
    {
        public void formatParagraph(List<string> textWithTags, Document doc)
        {
            object oCollapseEnd = WdCollapseDirection.wdCollapseEnd;
            doc.Paragraphs.Add();
            Range objRange = doc.Content;
            objRange.Collapse(ref oCollapseEnd);

            if (textWithTags[0].Contains("xmlabstr"))
            {
                var splitText = textWithTags[0].Split(']');
                for (int i = 0; i < splitText.Length; i++)
                {
                    if (!string.IsNullOrEmpty(splitText[i]))
                    {
                        objRange.Collapse(ref oCollapseEnd);

                        objRange.Font.Size = 11;
                        objRange.Paragraphs.Alignment = WdParagraphAlignment.wdAlignParagraphJustify;

                        if (i == 0)
                        {
                            objRange.Text = splitText[i] + "]";
                            objRange.Font.Name = "Arial";
                            objRange.Font.Color = WdColor.wdColorRed;
                        }
                        else if (i == 1)
                        {
                            objRange.Text = splitText[i] + "]";
                            objRange.Font.Name = "Times New Roman";
                            objRange.Font.Color = WdColor.wdColorBlue;
                        }
                        else
                        {
                            var rawTextSplit = splitText[i].Split('[');
                            objRange.Text = rawTextSplit[0];
                            objRange.Font.Name = "Times New Roman";
                            objRange.Font.Color = WdColor.wdColorBlack;

                            objRange.Collapse(ref oCollapseEnd);
                            objRange.Text = "[" + rawTextSplit[1] + "]";
                            objRange.Font.Name = "Arial";
                            objRange.Font.Color = WdColor.wdColorBlue;
                        }
                    }
                }
            }
        }
    }
}
