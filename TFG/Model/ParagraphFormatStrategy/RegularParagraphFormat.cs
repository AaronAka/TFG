using Microsoft.Office.Interop.Word;
using System.Collections.Generic;

namespace TFG.Model.ParagraphFormatStrategy
{
    public class RegularParagraphFormat : IParagraphFormat
    {
        public void formatParagraph(List<string> textWithTags, Document doc)
        {
            object oCollapseEnd = WdCollapseDirection.wdCollapseEnd;
            doc.Paragraphs.Add();
            Range objRange = doc.Content;
            objRange.Collapse(ref oCollapseEnd);

            objRange.Text = textWithTags[0];
            objRange.Font.Name = "Arial";
            objRange.Font.Size = 11;
            objRange.Paragraphs.Alignment = WdParagraphAlignment.wdAlignParagraphJustify;
        }
    }
}