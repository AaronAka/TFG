using Microsoft.Office.Interop.Word;
using System;
using System.Collections.Generic;
using System.Linq;

namespace TFG.Model
{
    public static class MarkDocumentFunctions
    {
        public static void AddSourceTags(List<string> extractedWords, Paragraph paragraph)
        {
            var organizations = paragraph.Range.Text.Split(";");
            foreach (var organization in organizations)
            {
                var organizationTrim = organization.TrimStart();
                var refNumber = organizationTrim.ElementAt(0);
                var indexNewOrg = extractedWords.Any() ? extractedWords.Count : 0;
                if (Char.IsDigit(refNumber))
                {
                    var organizationString = organizationTrim.Remove(0, 1);
                    //extractedWords.Add(string.Format(MarkingConstants.NORMAFF_OPEN, refNumber));
                    extractedWords.Add(MarkingConstants.SUP_OPEN + MarkingConstants.LABEL_OPEN);
                    extractedWords.Add(refNumber + "");
                    extractedWords.Add(MarkingConstants.SUP_OPEN + MarkingConstants.LABEL_CLOSE);
                    extractedWords.Add(string.Format(MarkingConstants.ORGDIVE_OPEN, refNumber));

                    var locationStart = organizationString.IndexOf(MarkingConstants.OPEN_PARENTHESIS);
                    var locationEnd = organizationString.IndexOf(MarkingConstants.CLOSE_PARENTHESIS);
                    var location = organizationString.Substring(locationStart + 1, locationEnd - locationStart - 1).Split(',');
                    var separateByBracket = organizationString.Split(MarkingConstants.OPEN_PARENTHESIS);

                    if (separateByBracket.Length > 0 && location.Length > 1)
                    {
                        extractedWords.Add(separateByBracket[0]);
                        extractedWords.Add(string.Format(MarkingConstants.ORGDIVE_CLOSE, refNumber));
                        extractedWords.Add(MarkingConstants.OPEN_PARENTHESIS);
                        extractedWords.Add(MarkingConstants.CITY_OPEN);
                        extractedWords.Add(location[0]);
                        extractedWords.Add(MarkingConstants.CITY_CLOSE);
                        extractedWords.Add(MarkingConstants.COMMA);
                        extractedWords.Add(MarkingConstants.COUNTRY_OPEN);
                        extractedWords.Add(location[1]);
                        extractedWords.Add(MarkingConstants.COUNTRY_CLOSE);
                        extractedWords.Add(MarkingConstants.CLOSE_PARENTHESIS + MarkingConstants.DOT);
                        extractedWords.Add(MarkingConstants.NORMAFF_CLOSE + "\r");

                        extractedWords.Insert(indexNewOrg, string.Format(MarkingConstants.NORMAFF_OPEN, refNumber, location[1]));
                    }
                }
            }
        }

        public static void AddKeywordTags(List<string> extractedWords, Paragraph paragraph)
        {
            var dividedKeywords = paragraph.Range.Text.Replace("\r", "").Split(':');
            var keywords = dividedKeywords[1].Split(';');

            extractedWords.Add(MarkingConstants.KWDGRP_OPEN);
            extractedWords.Add(MarkingConstants.SECTITLE_OPEN);
            extractedWords.Add(dividedKeywords[0] + MarkingConstants.COLON);
            extractedWords.Add(MarkingConstants.SECTITLE_CLOSE);

            foreach (string s in keywords)
            {
                extractedWords.Add(MarkingConstants.KWD_OPEN);
                extractedWords.Add(s);

                if (keywords.Last().Equals(s))
                {
                    extractedWords.Add(MarkingConstants.KWD_CLOSE);
                    extractedWords.Add(MarkingConstants.KWDGRP_CLOSE);
                }
                else
                {
                    extractedWords.Add(MarkingConstants.KWD_CLOSE + MarkingConstants.SEMICOLON);
                }
            }
        }

        public static void AddTableTags(int tableIndex, List<string> extractedWords, Paragraph paragraph)
        {
            var splitString = paragraph.Range.Text.Replace("\r", "").Replace("\u000e", "").Split('.');
            //taggedString = "[label]" + splitString[0] + "[/label].";
            //taggedString += "[caption]" + splitString[1] + splitString[splitString.Length - 1] + "[/caption]";
            extractedWords.Add(string.Format(MarkingConstants.FIGGRP_OPEN, tableIndex));
            extractedWords.Add(MarkingConstants.LABEL_OPEN);
            extractedWords.Add(splitString[0]);
            extractedWords.Add(MarkingConstants.LABEL_CLOSE);
            extractedWords.Add(MarkingConstants.CAPTION_OPEN);
            extractedWords.Add(splitString[1] + splitString[splitString.Length - 1]);
            extractedWords.Add(MarkingConstants.CAPTION_CLOSE);
        }

        public static void AddSecTypeTags(List<string> extractedWords, Paragraph paragraph, bool inBody)
        {
            var secType = GetSecType(paragraph.Range.Text);
            if (!inBody)
            {
                extractedWords.Add(string.Format(MarkingConstants.INITIAL_SECTYPE_OPEN, secType));
            }
            else
            {
                extractedWords.Add(string.Format(MarkingConstants.SECTYPE_OPEN, secType));
            }

            extractedWords.Add(MarkingConstants.SECTITLE_OPEN);
            extractedWords.Add(paragraph.Range.Text.Replace("\r", ""));
            extractedWords.Add(MarkingConstants.SECTITLE_CLOSE);
        }

        private static string GetSecType(string text)
        {
            if (text == "INTRODUCCIÓN\r")
            {
                return "intro";
            }
            else if (text == "MÉTODO\r")
            {
                return "methods";
            }
            else
            {
                return "results";
            }
        }

        public static void MarkAuthorsInterop(string par, List<string> extractedWords)
        {
            string[] splittedAuthors = par.Split(new char[] { ',', '.', 'y', '&' });
            string author = "";
            int i = 0;
            foreach (string val in splittedAuthors)
            {
                if (val.Trim().Length > 1 && !val.Any(char.IsDigit))
                {
                    author = string.Format(MarkingConstants.AUTHORS_OPEN, val);
                    i++;
                }
                else if (!val.Any(char.IsDigit))
                {
                    author += string.Format(MarkingConstants.AUTHORS_CLOSE, val);
                    extractedWords.Add(author + "\v");
                }
            }
        }

        public static void AddBibliographyTags(List<string> extractedWords, string[] splitText)
        {
            bool dateFound = false;
            bool additionalInformation = false;
            foreach (string line in splitText)
            {
                if (!dateFound && line.Contains(MarkingConstants.OPEN_PARENTHESIS))
                {
                    AddBaseBibliographyTags(extractedWords, splitText, line);
                    dateFound = true;
                    additionalInformation = !splitText.Last().Equals(line);
                }
                else if (additionalInformation)
                {
                    AddAdditionalInformation(extractedWords, line);
                }
                else if (line.Length > 3 && !line.Contains(MarkingConstants.DOT))
                {
                    //taggedString += " [pauthor][surname]" + line + "[/surname],";
                    extractedWords.Add(string.Format(MarkingConstants.BIBLIOGRAPHY_PAUTHOR_OPEN, line));
                }
                else
                {
                    //taggedString += " [fname]" + line + "[/fname][/pauthor],";
                    extractedWords.Add(string.Format(MarkingConstants.BIBLIOGRAPHY_PAUTHOR_CLOSE, line));
                }
            }
        }

        private static void AddAdditionalInformation(List<string> extractedWords, string line)
        {
            if (line.Contains(MarkingConstants.OPEN_PARENTHESIS))
            {
                // Publishing information

                var indexOfOpenBracket = line.IndexOf(MarkingConstants.OPEN_PARENTHESIS);
                var indexOfEndBracket = line.IndexOf(MarkingConstants.CLOSE_PARENTHESIS);
                //taggedString += string.Concat("[volid]", line.AsSpan(0, indexOfOpenBracket).Trim(), "[/volid]");
                extractedWords.Add(string.Format(MarkingConstants.BIBLIOGRAPHY_VOLID, line.Substring(0, indexOfOpenBracket).Trim()));
                //taggedString += string.Concat("([issueno]", line.AsSpan(indexOfOpenBracket + 1 , indexOfEndBracket - indexOfOpenBracket - 1).Trim(), "[/issueno]),");
                extractedWords.Add(string.Format(MarkingConstants.BIBLIOGRAPHY_ISSUENO, line.Substring(indexOfOpenBracket + 1, indexOfEndBracket - indexOfOpenBracket - 1).Trim()));
            }
            else if (line.Contains('–'))
            {
                int indexOfSeparation = line.IndexOf('.');

                if (indexOfSeparation > -1)
                {
                    // Page and article id

                    //taggedString += string.Concat("[pages]", line.AsSpan(0, indexOfSeparation).Trim(), "[/pages]");
                    extractedWords.Add(string.Format(MarkingConstants.BIBLIOGRAPHY_PAGES, line.Substring(0, indexOfSeparation).Trim()));
                    //taggedString += string.Concat("[pubid]", line.AsSpan(indexOfSeparation + 1).Trim() , "[/pubid]");
                    extractedWords.Add(string.Format(MarkingConstants.BIBLIOGRAPHY_PUBID, line.Substring(indexOfSeparation + 1).Trim()));
                }
                else
                {
                    // Only page information

                    //taggedString += "[pages]" + line.Trim() + "[/pages]";
                    extractedWords.Add(string.Format(MarkingConstants.BIBLIOGRAPHY_PAGES, line.Trim()));
                }
            }
            else if (line.Contains(':'))
            {
                // doi/publication url
                //taggedString += "[pubid]" + line.Trim() + "[/pubid]";
                extractedWords.Add(string.Format(MarkingConstants.BIBLIOGRAPHY_PUBID, line.Trim()));
            }
            else
            {
                // Only basic publishing information
                //taggedString += "[volid]" + line.Trim() + "[/volid]";
                extractedWords.Add(string.Format(MarkingConstants.BIBLIOGRAPHY_VOLID, line.Trim()));
            }
        }

        private static void AddBaseBibliographyTags(List<string> extractedWords, string[] splitText, string line)
        {
            var decomposedLastLine = line.Split(MarkingConstants.DOT);
            if (decomposedLastLine.Length >= 3)
            {
                var lastName = string.Empty;
                bool dateReached = false;
                int i = 0;
                while (!dateReached && i < decomposedLastLine.Length)
                {
                    if (!decomposedLastLine[i].Contains(MarkingConstants.OPEN_PARENTHESIS))
                    {
                        lastName += decomposedLastLine[i] + MarkingConstants.DOT;
                        i++;
                    }
                    else
                    {
                        dateReached = true;
                    }
                }

                //var date = decomposedLastLine[1].Remove('(').Remove(')');
                if (line == splitText.Last())
                {
                    //taggedString += " [fname]" + lastName + "[/fname][/pauthor]";
                    extractedWords.Add(string.Format(MarkingConstants.BIBLIOGRAPHY_PAUTHOR_CLOSE_NO_COMMA, lastName));
                }
                else
                {
                    //taggedString += " [fname]" + lastName + "[/fname][/pauthor], ";
                    extractedWords.Add(string.Format(MarkingConstants.BIBLIOGRAPHY_PAUTHOR_CLOSE, lastName));
                }

                var date = decomposedLastLine[i].Replace(MarkingConstants.OPEN_PARENTHESIS, "").Replace(MarkingConstants.CLOSE_PARENTHESIS, "").Trim();
                //taggedString += "[/authors] ";
                extractedWords.Add(MarkingConstants.BIBLIOGRAPHY_AUTHORS_CLOSE);
                //taggedString += "([date dateiso=\"" + (date + "0000") + "\" specyear=\"" + date + "\"]" + date + "[/date]).";
                extractedWords.Add(string.Format(MarkingConstants.BIBLIOGRAPHY_DATE, date));
                //taggedString += " [arttitle]" + decomposedLastLine[i + 1].Trim() + ".[/arttitle]";
                extractedWords.Add(string.Format(MarkingConstants.BIBLIOGRAPHY_ARTTITLE, decomposedLastLine[i + 1].Trim()));
                i++;

                if (i <= decomposedLastLine.Length - 2)
                {
                    //taggedString += "[source]" + decomposedLastLine[i + 1] + "[/source],";
                    extractedWords.Add(string.Format(MarkingConstants.BIBLIOGRAPHY_SOURCE, decomposedLastLine[i + 1]));
                }
            }
        }

        public static void AddDoctitleTags(List<string> extractedWords, Paragraph paragraph)
        {
            extractedWords.Clear();
            extractedWords.Add(MarkingConstants.DOCTITLE_OPEN);
            extractedWords.Add(paragraph.Range.Text.Replace("\r", "").Replace("\u000e", ""));
            extractedWords.Add(MarkingConstants.DOCTITLE_CLOSE);
        }
    }
}