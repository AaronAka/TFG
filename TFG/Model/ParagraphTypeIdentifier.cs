using Microsoft.Office.Interop.Word;

namespace TFG.Model
{
    public static class ParagraphTypeIdentifier
    {
        public static ParagraphType GetParagraphType(bool inBody, Paragraph paragraph)
        {
            if (EmptyParagraph(paragraph))
            {
                return ParagraphType.Empty;
            }
            else if (IsDoctitleParagraph(paragraph))
            {
                return ParagraphType.Doctitle;
            }
            else if (IsAuthorsParagraph(paragraph))
            {
                return ParagraphType.Authors;
            }
            else if (IsSectitleParagraph(paragraph))
            {
                return ParagraphType.Sectitle;
            }
            else if (IsSubsectionParagraph(paragraph))
            {
                return ParagraphType.Subsection;
            }
            else if (IsRegularParagraph(paragraph))
            {
                return ParagraphType.Regular;
            }
            else if (IsTableParagraph(paragraph))
            {
                return ParagraphType.Table;
            }
            else if (IsKeywordsParagraph(paragraph))
            {
                return ParagraphType.Keywords;
            }
            else if (IsSourcesParagraph(inBody, paragraph))
            {
                return ParagraphType.Sources;
            }
            else
            {
                return ParagraphType.Empty;
            }
        }

        private static bool IsSourcesParagraph(bool inBody, Paragraph paragraph)
        {
            return !inBody && paragraph.Range.Bold == 0 && paragraph.Range.Font.Size == 9999999 && paragraph.Alignment == WdParagraphAlignment.wdAlignParagraphCenter;
        }

        private static bool IsKeywordsParagraph(Paragraph paragraph)
        {
            return paragraph.Range.Bold == 9999999 && paragraph.Range.Font.Size == 11 && paragraph.Range.Text.Contains(';') && paragraph.Alignment == WdParagraphAlignment.wdAlignParagraphJustify;
        }

        private static bool IsTableParagraph(Paragraph paragraph)
        {
            return paragraph.Range.Text.Contains("Tabla") && paragraph.Range.Font.Size >= 8.5;
        }

        private static bool IsRegularParagraph(Paragraph paragraph)
        {
            return paragraph.Range.Text.Length > 2 && paragraph.Range.Font.Size >= 10 && paragraph.Range.Font.Size <= 11 && paragraph.Range.Bold != 9999999;
        }

        private static bool IsSubsectionParagraph(Paragraph paragraph)
        {
            return paragraph.Range.Case == WdCharacterCase.wdUpperCase && paragraph.Range.Bold == -1 && paragraph.Range.Font.Size == 11;
        }

        private static bool IsSectitleParagraph(Paragraph paragraph)
        {
            return paragraph.Range.Text == "INTRODUCCIÓN\r" || paragraph.Range.Text == "MÉTODO\r" || paragraph.Range.Text.Contains("RESULTADOS");
        }

        private static bool IsAuthorsParagraph(Paragraph paragraph)
        {
            return paragraph.Range.Bold == 0 && paragraph.Range.Italic == 0 && paragraph.Range.Font.Size == 9999999 && paragraph.Alignment == WdParagraphAlignment.wdAlignParagraphCenter;
        }

        private static bool IsDoctitleParagraph(Paragraph paragraph)
        {
            return paragraph.Range.Bold == -1 && paragraph.Range.Font.Size > 13;
        }

        public static bool EmptyParagraph(Paragraph paragraph)
        {
            return paragraph.Range.Text.Length <= 4;
        }
    }
}