using Microsoft.Office.Interop.Word;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TFG.Model
{
    public static class ParagraphTypeIdentifier
    {
        public static bool IsSourcesParagraph(bool inBody, Paragraph paragraph)
        {
            return !inBody && paragraph.Range.Bold == 0 && paragraph.Range.Font.Size == 9999999 && paragraph.Alignment == WdParagraphAlignment.wdAlignParagraphCenter;
        }

        public static bool IsKeywordsParagraph(Paragraph paragraph)
        {
            return paragraph.Range.Bold == 9999999 && paragraph.Range.Font.Size == 11 && paragraph.Range.Text.Contains(';') && paragraph.Alignment == WdParagraphAlignment.wdAlignParagraphJustify;
        }

        public static bool IsTableParagraph(Paragraph paragraph)
        {
            return paragraph.Range.Text.Contains("Tabla") && paragraph.Range.Font.Size >= 8.5;
        }

        public static bool IsRegularParagraph(Paragraph paragraph)
        {
            return paragraph.Range.Text.Length > 2 && paragraph.Range.Font.Size >= 10 && paragraph.Range.Font.Size <= 11 && paragraph.Range.Bold != 9999999;
        }

        public static bool IsSubsectionParagraph(Paragraph paragraph)
        {
            return paragraph.Range.Case == WdCharacterCase.wdUpperCase && paragraph.Range.Bold == -1 && paragraph.Range.Font.Size == 11;
        }

        public static bool IsSectitleParagraph(Paragraph paragraph)
        {
            return paragraph.Range.Text == "INTRODUCCIÓN\r" || paragraph.Range.Text == "MÉTODO\r" || paragraph.Range.Text.Contains("RESULTADOS");
        }

        public static bool IsAuthorsParagraph(Paragraph paragraph)
        {
            return paragraph.Range.Bold == 0 && paragraph.Range.Italic == 0 && paragraph.Range.Font.Size == 9999999 && paragraph.Alignment == WdParagraphAlignment.wdAlignParagraphCenter;
        }

        public static bool IsDoctitleParagraph(Paragraph paragraph)
        {
            return paragraph.Range.Bold == -1 && paragraph.Range.Font.Size > 13;
        }

        public static bool EmptyParagraph(Paragraph paragraph)
        {
            return paragraph.Range.Text.Length <= 4;
        }
    }
}
