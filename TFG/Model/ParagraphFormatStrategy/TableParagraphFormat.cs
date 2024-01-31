using Microsoft.Office.Interop.Word;
using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using Range = Microsoft.Office.Interop.Word.Range;

namespace TFG.Model.ParagraphFormatStrategy
{
    public class TableParagraphFormat : IParagraphFormat
    {
        public void formatParagraph(List<string> textWithTags, Document doc)
        {
            object oCollapseEnd = WdCollapseDirection.wdCollapseEnd;
            doc.Paragraphs.Add();
            Range objRange = doc.Content;
            objRange.Collapse(ref oCollapseEnd);

            for (int i = 0; i < textWithTags.Count; i++)
            {
                objRange.Collapse(ref oCollapseEnd);

                objRange.Text = textWithTags[i];
                objRange.Font.Name = "Arial";
                objRange.Font.Size = 9;
                objRange.Paragraphs.Alignment = WdParagraphAlignment.wdAlignParagraphJustify;

                if (textWithTags[i].Contains("figgrp"))
                {
                    objRange.Font.Color = WdColor.wdColorAqua;
                }
                else if (textWithTags[i].Contains("label"))
                {
                    objRange.Font.Color = WdColor.wdColorRed;
                }
                else if (textWithTags[i].Contains("caption"))
                {

                    objRange.Font.Color = WdColor.wdColorGreen;
                }
                else
                {
                    objRange.Font.Name = "Times New Roman";
                    objRange.Font.Color = WdColor.wdColorBlack;
                }
            }

            objRange.Collapse(ref oCollapseEnd);
            objRange.Text = "\r[graphic href=1578-8423-CPD-22-1-1-13]";
            objRange.Font.Name = "Arial";
            objRange.Font.Size = 9;
            objRange.Font.Color = WdColor.wdColorRose;
            objRange.Paragraphs.Alignment = WdParagraphAlignment.wdAlignParagraphJustify;

            Match match = Regex.Match(textWithTags[0], @"\d+");

            objRange.Collapse(ref oCollapseEnd);
            objRange.Text = "[/graphic]";
            objRange.Font.Name = "Arial";
            objRange.Font.Size = 9;
            objRange.Font.Color = WdColor.wdColorRose;
            objRange.InlineShapes.AddPicture("C:\\Users\\PC\\Documents\\TFG\\1578-8423-CPD-22-1-1-13image00" + Convert.ToInt32(match.Value) + ".png");
            objRange.Paragraphs.Alignment = WdParagraphAlignment.wdAlignParagraphJustify;

            objRange.Collapse(ref oCollapseEnd);
            objRange.Text = "[/figgrp]";
            objRange.Font.Name = "Arial";
            objRange.Font.Size = 9;
            objRange.Font.Color = WdColor.wdColorAqua;
            objRange.Paragraphs.Alignment = WdParagraphAlignment.wdAlignParagraphJustify;
        }
    }
}
