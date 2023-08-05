using Microsoft.Office.Interop.Word;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TFG.Model
{
    public class MarkDocumentFunctions
    {
        public void MarkDocument (Document importedDoc, Document newDoc)
        {
            var a = importedDoc.Paragraphs;
            var b = newDoc.Paragraphs;

        }
    }
}
