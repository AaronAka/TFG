using DocumentFormat.OpenXml.ExtendedProperties;
using Microsoft.Office.Interop.Word;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TFG.Model.ParagraphFormatStrategy
{
    public class RegularParagraphFormat : IParagraphFormat
    {
        public void formatParagraph(List<string> textWithTags, Document doc)
        {
            object oCollapseEnd = WdCollapseDirection.wdCollapseEnd;
            doc.Paragraphs.Add();
            Microsoft.Office.Interop.Word.Range objRange = doc.Content;
            objRange.Collapse(ref oCollapseEnd);

            objRange.Text = textWithTags[0];
            objRange.Font.Name = "Arial";
            objRange.Font.Size = 11;
            objRange.Paragraphs.Alignment = WdParagraphAlignment.wdAlignParagraphJustify;
        }
    }
}
