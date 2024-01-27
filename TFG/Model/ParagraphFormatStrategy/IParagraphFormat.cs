using Microsoft.Office.Interop.Word;
using System.Collections.Generic;

namespace TFG.Model.ParagraphFormatStrategy
{
    public interface IParagraphFormat
    {
        void formatParagraph(List<string> textWithTags, Document doc);
    }
}
