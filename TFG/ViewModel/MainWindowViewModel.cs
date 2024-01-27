using CommunityToolkit.Mvvm.Input;
using Microsoft.Win32;
using Microsoft.Office.Interop.Word;
using System;
using System.Windows;
using System.Windows.Input;
using System.ComponentModel;
using Word = Microsoft.Office.Interop.Word;
using Xceed.Words.NET;
using System.Linq;
using Xceed.Document.NET;
using Document = Microsoft.Office.Interop.Word.Document;
using System.Collections.Generic;
using System.Drawing;
using System.Xml.Linq;
using Application = Microsoft.Office.Interop.Word.Application;
using System.Reflection;
using DocumentFormat.OpenXml.Math;
using Paragraph = Microsoft.Office.Interop.Word.Paragraph;
using DocumentFormat.OpenXml.Drawing.Charts;
using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Office2010.Excel;
using DocumentFormat.OpenXml.Bibliography;
using System.Runtime.InteropServices.ObjectiveC;
using Shape = Microsoft.Office.Interop.Word.Shape;
using TFG.Model.ParagraphFormatStrategy;

namespace TFG.ViewModel
{
    public class MainWindowViewModel : INotifyPropertyChanged
    {
        private ICommand _openFileDialogCommand;
        private string _fileContent;
        private bool _enabledMarkedButton;
        private ICommand _markFileCommand;
        private string _importedDocument;
        private string _xmlFormat;
        private IParagraphFormat strategy;

        public MainWindowViewModel() 
        {
            _openFileDialogCommand = new RelayCommand(ReadUserSelectedFile, ReadUserSelectedFileCanExecute);
            _markFileCommand = new RelayCommand(MarkImportedDocument, ReadUserSelectedFileCanExecute);
            _fileContent = string.Empty;
            _importedDocument = string.Empty;
        }
        
        public ICommand OpenFileDialogCommand
        {
            get
            {
                return _openFileDialogCommand;
            }
            set
            {
                if (value != null)
                {
                    _openFileDialogCommand = value;
                }
            }
        }

        public ICommand MarkFileCommand
        {
            get
            {
                return _markFileCommand;
            }
            set
            {
                if (value != null)
                {
                    _markFileCommand = value;
                }
            }
        }

        public string FileContent
        {
            get
            {
                return _fileContent;
            }
            set
            {
                if (!string.IsNullOrEmpty(value))
                {
                    _fileContent = value;
                }
                OnPropertyChanged(nameof(FileContent));
            }
        }

        public bool EnabledMarkButton
        {
            get
            {
                return _enabledMarkedButton;
            }
            set
            {
                if (value != _enabledMarkedButton)
                {
                    _enabledMarkedButton = value;
                }
                OnPropertyChanged(nameof(EnabledMarkButton));
            }
        }

        private bool ReadUserSelectedFileCanExecute()
        {
            return true;
        }

        private void SetStrategy(IParagraphFormat strategy)
        {
            this.strategy = strategy;
        }

        private void ReadUserSelectedFile()
        {
            try
            {
                OpenFileDialog fileDialog = new()
                {
                    Filter = "Word documents (.doc; .docx)|*.doc;*.docx;*.pdf"
                };

                fileDialog.ShowDialog();

                string selectedFilePath  = fileDialog.FileName.ToString();

                if (!string.IsNullOrEmpty(selectedFilePath))
                {
                    Object miss = Type.Missing;
                    object readOnly = true;
                    object isVisible = false;

                    Document document;
                    Word.Application application = new() { Visible = false };
                    _importedDocument = selectedFilePath;

                    document = application.Documents.Open(selectedFilePath, ref miss, ref readOnly, ref miss, ref miss, ref miss, ref miss, ref miss, 
                                                            ref miss, ref miss, ref miss, ref isVisible, ref miss, ref miss, ref miss, ref miss);
                    document.ActiveWindow.Selection.WholeStory();
                    document.ActiveWindow.Selection.Copy();

                    var aa = document.XMLNodes;
                    _xmlFormat = document.WordOpenXML;
                    var dataDoc = "";
                    dataDoc = Clipboard.GetDataObject().GetData(DataFormats.Rtf).ToString();
                    
                    if(dataDoc != null)
                    {
                        FileContent = dataDoc;
                        EnabledMarkButton = true;
                    }

                    application.Quit();
                }
            } 
            catch (Exception ex)
            {
                MessageBox.Show("An error has occurred while importing the file " + ex.ToString());
            }
        }

        private void MarkImportedDocument()
        {
            if( _importedDocument != null )
            {
                Object miss = Type.Missing;
                object readOnly = true;
                object isVisible = false;

                Application application = new Application();
                Document document = application.Documents.Add(_importedDocument);

                //Application application2 = new Application();
                Document newDocument = application.Documents.Add();
                /*newDocument.PageSetup.TopMargin = document.PageSetup.TopMargin;
                newDocument.PageSetup.LeftMargin = document.PageSetup.LeftMargin;
                newDocument.PageSetup.RightMargin = document.PageSetup.RightMargin;
                newDocument.PageSetup.BottomMargin = document.PageSetup.BottomMargin;*/

                Find findObject = application.Selection.Find;

                var tableIndex = 1;
                var refIndex = 1;
                Boolean inBody = false;
                Boolean reachedBibliography = false;
                bool inFirstSection = false;
                bool inResults = false;
                foreach (Word.Paragraph paragraph in document.Paragraphs)
                {
                    string taggedString;
                    if (paragraph.Range.Text.Length <= 4)
                    {
                        continue;
                    }
                    /*if (paragraph.Range.Text.Contains('\u000e'))
                    {
                       
                    }*/
                    if (paragraph.Range.Bold == -1 && paragraph.Range.Font.Size > 13)
                    {
                        //FindAndReplaceText(paragraph.Range.Text, findObject);
                        taggedString = "[doctile]" + paragraph.Range.Text.Replace("\r", "").Replace("\u000e", "") + "[/doctile]";
                        //taggedString = paragraph.Range.Text.Replace("\r", "").Replace("\u000e", "");
                        BuildDoctitleString(taggedString, newDocument);
                        //InsertTextInNewDoc(taggedString, newDocument);
                    }
                    else if (reachedBibliography)
                    {
                        taggedString = "[ref id=\"r" + refIndex + "\" reftype =\"journal\"][authors role=\"nd\"]";
                        var text = paragraph.Range.Text.Replace("\r", "").Replace("\u000e", "");
                        Boolean additionalInformation = false;
                        Boolean dateFound = false;
                        var splitText = text.Split(',');
                        
                        foreach (string line in splitText)
                        {
                            if (!dateFound && line.Contains('('))
                            {
                                var c = 0;
                                var decomposedLastLine = line.Split('.');
                                if (decomposedLastLine.Length >= 3)
                                {
                                    var lastName = string.Empty;
                                    bool dateReached = false;
                                    int i = 0;
                                    while (!dateReached && i < decomposedLastLine.Length)
                                    {
                                        if (!decomposedLastLine[i].Contains('('))
                                        {
                                            lastName += decomposedLastLine[i] + ".";
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
                                        taggedString += " [fname]" + lastName + "[/fname][/pauthor]";
                                    }
                                    else
                                    {
                                        taggedString += " [fname]" + lastName + "[/fname][/pauthor], ";
                                    }
                                    
                                    var date = decomposedLastLine[i].Replace("(", "").Replace(")", "").Trim();
                                    taggedString += "[/authors] ";
                                    taggedString += "([date dateiso=\"" + (date + "0000") + "\" specyear=\"" + date + "\"]" + date + "[/date]).";
                                    taggedString += " [arttitle]" + decomposedLastLine[i + 1].Trim() + ".[/arttitle]";
                                    i++;

                                    if (i <= decomposedLastLine.Length - 2)
                                    {
                                        taggedString += "[source]" + decomposedLastLine[i + 1] + "[/source],";
                                    }

                                    dateFound = true;
                                    additionalInformation = !splitText.Last().Equals(line);
                                }
                            }
                            else if (additionalInformation)
                            {
                                var dd = 0;
                                if (line.Contains('('))
                                {
                                    var indexOfOpenBracket = line.IndexOf("(");
                                    var indexOfEndBracket = line.IndexOf(")");
                                    taggedString += string.Concat("[volid]", line.AsSpan(0, indexOfOpenBracket).Trim(), "[/volid]");
                                    taggedString += string.Concat("([issueno]", line.AsSpan(indexOfOpenBracket + 1 , indexOfEndBracket - indexOfOpenBracket - 1).Trim(), "[/issueno]),");
                                }
                                else if (line.Contains('–'))
                                {
                                    int indexOfSeparation = line.IndexOf('.');

                                    if (indexOfSeparation > -1)
                                    {
                                        taggedString += string.Concat("[pages]", line.AsSpan(0, indexOfSeparation).Trim(), "[/pages]");
                                        taggedString += string.Concat("[pubid]", line.AsSpan(indexOfSeparation + 1).Trim() , "[/pubid]");
                                    } 
                                    else
                                    {
                                        taggedString += "[pages]" + line.Trim() + "[/pages]";
                                    }
                                }
                                else if (line.Contains(':'))
                                {
                                    taggedString += "[pubid]" + line.Trim() + "[/pubid]";
                                }
                                else
                                {
                                    taggedString += "[volid]" + line.Trim() + "[/volid]";
                                }
                            }
                            else if (line.Length > 3 && !line.Contains('.'))
                            {
                                taggedString += " [pauthor][surname]" + line + "[/surname],";
                            }
                            else
                            {
                                taggedString += " [fname]" + line + "[/fname][/pauthor],";
                            }
                        }

                        taggedString += "[/ref]";

                        refIndex++;
                        BuildBibliographyParagraphString(taggedString, newDocument);
                    }
                    else if (paragraph.Range.Bold == 0 && paragraph.Range.Italic == 0 && paragraph.Range.Font.Size == 9999999 && paragraph.Alignment == WdParagraphAlignment.wdAlignParagraphCenter)
                    {
                        taggedString = MarkAuthorsInterop(paragraph.Range.Text);
                        BuildAuthorsString(taggedString, newDocument);
                    }
                    else if (paragraph.Range.Text == "INTRODUCCIÓN\r")
                    {
                        //taggedString = "[xmlbody]\r [sec sec-type=\"intro\"][sectitle]" + paragraph.Range.Text.Replace("\r", "") + "[/sectitle]";
                        inBody = true;
                        List<string> strings = new List<string>();
                        strings.Add("[xmlbody]\r [sec sec-type=\"intro\"]");
                        strings.Add("[sectitle]");
                        strings.Add(paragraph.Range.Text.Replace("\r", ""));
                        strings.Add("[/sectitle]");
                        BuildIntroParagraphString(strings, newDocument);
                    }
                    else if (paragraph.Range.Text == "MÉTODO\r")
                    {
                        //taggedString = "[sec sec-type=\"methods\"][sectitle]" + paragraph.Range.Text.Replace("\r", "") + "[/sectitle]";
                        inFirstSection = true;
                        List<string> strings = new List<string>();
                        strings.Add("[sec sec-type=\"methods\"]");
                        strings.Add("[sectitle]");
                        strings.Add(paragraph.Range.Text.Replace("\r", ""));
                        strings.Add("[/sectitle]");
                        BuildIntroParagraphString(strings, newDocument);
                        //BuildRegularParagraphString(taggedString, newDocument);
                    }
                    else if (paragraph.Range.Text.Contains("RESULTADOS"))
                    {
                        //taggedString = "[sec sec-type=\"results\"][sectitle]" + paragraph.Range.Text.Replace("\r", "") + "[/sectitle]";
                        //BuildRegularParagraphString(taggedString, newDocument);
                        List<string> strings = new List<string>();
                        strings.Add("[sec sec-type=\"results\"]");
                        strings.Add("[sectitle]");
                        strings.Add(paragraph.Range.Text.Replace("\r", ""));
                        strings.Add("[/sectitle]");
                        inResults = true;
                        BuildIntroParagraphString(strings, newDocument);
                    }
                    else if (paragraph.Range.Case == WdCharacterCase.wdUpperCase && paragraph.Range.Bold == -1 && paragraph.Range.Font.Size == 11)
                    {
                        if (!inBody)
                        {
                            taggedString = "[xmlabstr language=\"es\"][sectitle]" + paragraph.Range.Text.Replace("\r", "") + "[/sectitle]";
                            BuildSectitleString(taggedString, newDocument);
                        }
                        else
                        {
                            taggedString = "[subsec][sectitle]" + paragraph.Range.Text.Replace("\r", "") + "[/sectitle]";
                            if (taggedString.Contains("BIBLIOGRAFÍA"))
                            {
                                reachedBibliography = true;
                            }
                            BuildRegularParagraphString(taggedString, newDocument);
                        }
                    }
                    else if (paragraph.Range.Text.Length > 2 && paragraph.Range.Font.Size >= 10 && paragraph.Range.Font.Size <= 11 && paragraph.Range.Bold != 9999999)
                    {
                        List<string> strings = new List<string>();
                        var rawText = paragraph.Range.Text.Replace("\r", "").Replace("\u000e", "");
                        if (inFirstSection || inResults)
                        {
                            strings.Add("[p]" + rawText + "[/p]");

                            SetStrategy(new RegularParagraphFormat());
                            strategy.formatParagraph(strings, newDocument);
                        }
                        else
                        {
                            strings.Add("[p]");
                            strings.Add(rawText);
                            strings.Add("[/p]");

                            SetStrategy(new RegularParagraphColoured());
                            strategy.formatParagraph(strings, newDocument);
                        }
                    }
                    else if (paragraph.Range.Text.Contains("Tabla")  && paragraph.Range.Font.Size >= 8.5)
                    {
                        var splitString = paragraph.Range.Text.Replace("\r", "").Replace("\u000e", "").Split('.');
                        //taggedString = "[label]" + splitString[0] + "[/label].";
                        //taggedString += "[caption]" + splitString[1] + splitString[splitString.Length - 1] + "[/caption]";
                        List<string> strings = new List<string>();
                        strings.Add("[figgrp id =" + tableIndex + "]");
                        strings.Add("[label]");
                        strings.Add(splitString[0]);
                        strings.Add("[/label]");
                        strings.Add("[caption]");
                        strings.Add(splitString[1] + splitString[splitString.Length - 1]);
                        strings.Add("[/caption]");

                        BuildTableWithString(strings, tableIndex , newDocument);
                        tableIndex++;

                    }
                    else if (paragraph.Range.Bold == 9999999 && paragraph.Range.Font.Size == 11 && paragraph.Range.Text.Contains(';') && paragraph.Alignment == WdParagraphAlignment.wdAlignParagraphJustify)
                    {
                        var dividedKeywords = paragraph.Range.Text.Replace("\r", "").Split(':');
                        var keywords = dividedKeywords[1].Split(';');
                        List<string> strings = new List<string>();
                        strings.Add("[kwdgrp language=\"es\"]");
                        strings.Add("[sectitle]");
                        strings.Add(dividedKeywords[0] + ":");
                        strings.Add("[/sectitle]");

                        foreach(string s in keywords)
                        {
                            if (keywords.Last().Equals(s))
                            {
                                strings.Add("[kwd]");
                                strings.Add(s);
                                strings.Add("[/kwd]");
                                strings.Add("[/kwdgrp]");
                            }
                            else
                            {
                                strings.Add("[kwd]");
                                strings.Add(s);
                                strings.Add("[/kwd];");
                            }
                        }

                        BuildKeywordParagraph(strings, newDocument);

                        /*taggedString = "[kwdgrp language=\"es\"][sectitle]" + dividedKeywords[0] + ":" + "[/sectitle]";
                        var keywords = dividedKeywords[1].Split(';');
                        foreach (var word in keywords)
                        {
                            if (keywords.Last().Equals(word))
                            {
                                taggedString += "[kwd]" + word + "[/kwd][/kwdgrp]";
                            }
                            else
                            {
                                taggedString += "[kwd]" + word + "[/kwd];";
                            }

                        }
                        BuildRegularParagraphString(taggedString, newDocument );*/
                    }
                    else if (!inBody && paragraph.Range.Bold == 0 && paragraph.Range.Font.Size == 9999999 && paragraph.Alignment == WdParagraphAlignment.wdAlignParagraphCenter)
                    { 
                        var organizations = paragraph.Range.Text.Split(";");
                        List<string> strings;
                        foreach (var organization in organizations)
                        {
                            var refNumber = organization.Trim().ElementAt(0);
                            if (Char.IsDigit(refNumber))
                            {
                                strings = new List<string>();
                                var organizationString = organization.Remove(0, 1);
                                strings.Add("[normaff id =\"aff" + refNumber + "\" ncountry=\"Spain\" ]");
                                strings.Add("[label][sup]");
                                strings.Add(refNumber + "");
                                strings.Add("[/sup][/label]");
                                strings.Add("[orgdiv" + refNumber + "]");

                                var locationStart = organizationString.IndexOf('(');
                                var locationEnd = organizationString.IndexOf(')');
                                var location = organizationString.Substring(locationStart + 1, locationEnd - locationStart - 1).Split(',');
                                var separateByBracket = organizationString.Split('(');

                                if (separateByBracket.Length > 0 && location.Length > 1)
                                {
                                    strings.Add(separateByBracket[0]);
                                    strings.Add("[/orgdiv" + refNumber + "]");
                                    strings.Add("(");
                                    strings.Add("[city]");
                                    strings.Add(location[0]);
                                    strings.Add("[/city]");
                                    strings.Add(",");
                                    strings.Add("[country]");
                                    strings.Add(location[1]);
                                    strings.Add(").");
                                    strings.Add("[normaff]\r");
                                }

                                BuildSourcesString(strings, newDocument);
                                /* var organizationString = organization.Remove(0, 1);
                                taggedString = "[normaff id=\"aff" + refNumber + "\" ncountry=\"Spain\" ] [label][sup]" + refNumber + "[/sup][/label][orgdiv" + refNumber + "]";
                                var locationStart = organizationString.IndexOf('(');
                                var locationEnd = organizationString.IndexOf(')');
                                var location = organizationString.Substring(locationStart + 1, locationEnd - locationStart - 1).Split(',');
                                var separateByBracket = organizationString.Split('(');

                                if (separateByBracket.Length > 0 && location.Length > 1)
                                {
                                    taggedString += separateByBracket[0];
                                    taggedString += "[/orgdiv" + refNumber + "]([city]" + location[0] + "[/city],[country]" + location[1] + "[/country]).[normaff]\r";
                                }

                                BuildSourcesString(taggedString, newDocument);*/
                            }
                        }
                    }
                }

                object fileName = "C:\\Users\\PC\\Documents\\TFG\\pruebas\\prueba.docx";

                newDocument.PageSetup.TopMargin = application.InchesToPoints((float)0.5);
                newDocument.PageSetup.BottomMargin = application.InchesToPoints((float)0.5);
                newDocument.PageSetup.LeftMargin = application.InchesToPoints((float)0.5);
                newDocument.PageSetup.RightMargin = application.InchesToPoints((float)0.5);
                application.DisplayAlerts = WdAlertLevel.wdAlertsNone;
                newDocument.SaveAs2(ref fileName);
                document.Close(ref miss, ref miss, ref miss);
                document = null;
                application.Quit(ref miss, ref miss, ref miss);
                application = null;
                /*document.SaveAs2(ref fileName);
                document.Close(ref miss, ref miss, ref miss);
                document = null;
                application.Quit(ref miss, ref miss, ref miss);
                application = null;*/
            }
        }

        private void FindAndReplaceText(string replacingText,Find findObject)
        {
            Object miss = Type.Missing;

            findObject.ClearFormatting();
            findObject.Text = replacingText;
            findObject.Replacement.Text = "Fortnite ou llea\r";

            object replaceAll = WdReplace.wdReplaceAll;
            findObject.Execute(ref miss, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss, ref replaceAll, ref miss, ref miss, ref miss, ref miss);

        }

        private void BuildDoctitleString(string rawText, Document doc)
        {
            object oCollapseEnd = Word.WdCollapseDirection.wdCollapseEnd;
            doc.Paragraphs.Add();
            Word.Range objRange = doc.Content;
            var tags = "[doctile][/doctile]";
            objRange.Collapse(ref oCollapseEnd);

            objRange.Text = rawText;
            objRange.Bold = 0;
            objRange.Font.Name = "Times New Roman";
            objRange.Font.Size = 16;
            objRange.Paragraphs.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;

            for (int i = 1; i <= objRange.Words.Count; i++)
            {
                if (tags.Contains(objRange.Words[i].Text))
                {
                    objRange.Words[i].Font.Color = WdColor.wdColorPlum;
                }
            }

            //objRange.Text = rawText;
            /*objRange.Text = "[doctile]";
            objRange.Bold = 0;
            objRange.Font.Name = "Times New Roman";
            objRange.Font.Size = 16;
            objRange.Paragraphs.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
            objRange.Font.Color = WdColor.wdColorPlum;

            objRange.Collapse();
            objRange.Text = rawText;
            objRange.Bold = 0;
            objRange.Font.Name = "Times New Roman";
            objRange.Font.Size = 16;
            objRange.Paragraphs.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
            objRange.Font.Color = WdColor.wdColorBlack;

            objRange.Collapse();
            objRange.Text = "[/doctile]";
            objRange.Bold = 0;
            objRange.Font.Name = "Times New Roman";
            objRange.Font.Size = 16;
            objRange.Paragraphs.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
            objRange.Font.Color = WdColor.wdColorPlum;*/
        }

        private void BuildSourcesString(string rawText, Document doc)
        {
            object oCollapseEnd = Word.WdCollapseDirection.wdCollapseEnd;
            doc.Paragraphs.Add();
            Word.Range objRange = doc.Content;
            objRange.Collapse(ref oCollapseEnd);
            var tagGreen = "normaff aff ncountry norgname icountry city country";
            var tagRed = "label sup";
            var tagBlue = "orgdiv";
            var selectedColour = WdColor.wdColorGreen;

            objRange.Text = rawText;
            objRange.Bold = 0;
            objRange.Font.Name = "Times New Roman";
            objRange.Font.Size = 11;
            objRange.Paragraphs.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;

            for (int i = 1; i <= objRange.Words.Count; i++)
            {
                if (i < objRange.Words.Count)
                {
                    var word = objRange.Words[i].Text;
                    if (tagGreen.Contains(word))
                    {
                        selectedColour = WdColor.wdColorGreen;
                    }
                    else if (tagRed.Contains(word))
                    {
                        selectedColour = WdColor.wdColorRed;
                    }
                    else if (tagBlue.Contains(word))
                    {
                        selectedColour= WdColor.wdColorBlue;
                    }

                    objRange.Words[i].Font.Color = selectedColour;
                }

                /*if (i == objRange.Words.Count)
                {
                    objRange.Words[i - 1].Font.Color = objRange.Words[i].Font.Color;
                }*/
            }
        }

        private void BuildSourcesString(List<string> strings, Document doc)
        {
            object oCollapseEnd = Word.WdCollapseDirection.wdCollapseEnd;
            doc.Paragraphs.Add();
            Word.Range objRange = doc.Content;

            for (int i = 0; i < strings.Count; i++)
            {
                objRange.Collapse(ref oCollapseEnd);
                
                if (i == 0 || i == 8 || i == 10 || i == 12 || i == 15)
                {
                    objRange.Text = strings[i];
                    objRange.Bold = 0;
                    objRange.Font.Name = "Arial";
                    objRange.Font.Size = 12;
                    objRange.Font.Color = WdColor.wdColorGreen;
                    objRange.Paragraphs.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
                }
                else if (i == 1 || i == 3)
                {
                    objRange.Text = strings[i];
                    objRange.Bold = 0;
                    objRange.Font.Name = "Arial";
                    objRange.Font.Size = 12;
                    objRange.Font.Color = WdColor.wdColorRed;
                    objRange.Paragraphs.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
                }
                else if (i == 2 || i == 5 || i == 7 || i == 9 || i == 11 || i == 13 || i == 14)
                {
                    objRange.Text = strings[i];
                    objRange.Bold = 0;
                    objRange.Font.Name = "Times New Roman";
                    objRange.Font.Size = 12;
                    objRange.Font.Color = WdColor.wdColorBlack;
                    objRange.Paragraphs.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
                }
                else if (i == 4 || i == 6)
                {
                    objRange.Text = strings[i];
                    objRange.Bold = 0;
                    objRange.Font.Name = "Arial";
                    objRange.Font.Size = 12;
                    objRange.Font.Color = WdColor.wdColorBlue;
                    objRange.Paragraphs.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
                }

            }
        }


        private void BuildAuthorsString(string rawText, Document doc)
        {
            object oCollapseEnd = Word.WdCollapseDirection.wdCollapseEnd;
            doc.Paragraphs.Add();
            Word.Range objRange = doc.Content;
            objRange.Collapse(ref oCollapseEnd);

            objRange.Text = rawText;
            objRange.Font.Name = "Times New Roman";
            objRange.Font.Size = 12;
            objRange.Paragraphs.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
        }

        private void BuildSectitleString(string rawText, Document doc)
        {
            object oCollapseEnd = Word.WdCollapseDirection.wdCollapseEnd;
            doc.Paragraphs.Add();
            Word.Range objRange = doc.Content;
            objRange.Collapse(ref oCollapseEnd);

            if (rawText.Contains("xmlabstr"))
            {
                var splitText = rawText.Split(']');
                for (int i = 0; i < splitText.Count(); i++) 
                {
                    if (!string.IsNullOrEmpty(splitText[i]))
                    {
                        objRange.Collapse(ref oCollapseEnd);
                        if (i == 0)
                        {
                            objRange.Text = splitText[i] + "]";
                            objRange.Font.Name = "Arial";
                            objRange.Font.Size = 11;
                            objRange.Font.Color = WdColor.wdColorRed;
                            objRange.Paragraphs.Alignment = WdParagraphAlignment.wdAlignParagraphJustify;
                        }
                        else if (i == 1)
                        {
                            objRange.Text = splitText[i] + "]";
                            objRange.Font.Name = "Times New Roman";
                            objRange.Font.Size = 11;
                            objRange.Font.Color = WdColor.wdColorBlue;
                            objRange.Paragraphs.Alignment = WdParagraphAlignment.wdAlignParagraphJustify;
                        }
                        else
                        {
                            var rawTextSplit = splitText[i].Split('[');
                            objRange.Text = rawTextSplit[0];
                            objRange.Font.Name = "Times New Roman";
                            objRange.Font.Size = 11;
                            objRange.Font.Color = WdColor.wdColorBlack;
                            objRange.Paragraphs.Alignment = WdParagraphAlignment.wdAlignParagraphJustify;

                            objRange.Collapse(ref oCollapseEnd);
                            objRange.Text = "[" + rawTextSplit[1] + "]";
                            objRange.Font.Name = "Arial";
                            objRange.Font.Size = 11;
                            objRange.Font.Color = WdColor.wdColorBlue;
                            objRange.Paragraphs.Alignment = WdParagraphAlignment.wdAlignParagraphJustify;
                        }
                    }
                }
            }

            /*objRange.Text = rawText;
            objRange.Font.Name = "Times New Roman";
            objRange.Font.Size = 11;
            objRange.Paragraphs.Alignment = WdParagraphAlignment.wdAlignParagraphJustify;*/

        }

        private void BuildIntroParagraphString(List<string> rawText, Document doc)
        {
            object oCollapseEnd = Word.WdCollapseDirection.wdCollapseEnd;
            doc.Paragraphs.Add();
            Word.Range objRange = doc.Content;

            for(int i = 0; i < rawText.Count; i++)
            {
                objRange.Collapse(ref oCollapseEnd);
                if(i == 0)
                {
                    objRange.Text = rawText[i];
                    objRange.Font.Name = "Arial";
                    objRange.Font.Size = 11;
                    objRange.Font.Color = WdColor.wdColorGreen;
                    objRange.Paragraphs.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
                }
                else if (i == 1 || i == 3)
                {
                    objRange.Text = rawText[i];
                    objRange.Font.Name = "Arial";
                    objRange.Font.Size = 11;
                    objRange.Font.Color = WdColor.wdColorBlue;
                    objRange.Paragraphs.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
                }
                else
                {
                    objRange.Text = rawText[i];
                    objRange.Font.Name = "Times New Roman";
                    objRange.Font.Size = 11;
                    objRange.Font.Color = WdColor.wdColorBlack;
                    objRange.Paragraphs.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
                }
            }

            /*objRange.Text = rawText;
            objRange.Font.Name = "Times New Roman";
            objRange.Font.Size = 11;
            objRange.Paragraphs.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;*/
        }

        private void BuildRegularParagraphString(string rawText, Document doc)
        {
            object oCollapseEnd = Word.WdCollapseDirection.wdCollapseEnd;
            doc.Paragraphs.Add();
            Word.Range objRange = doc.Content;
            objRange.Collapse(ref oCollapseEnd);

            objRange.Text = rawText;
            objRange.Font.Name = "Arial";
            objRange.Font.Size = 11;
            objRange.Paragraphs.Alignment = WdParagraphAlignment.wdAlignParagraphJustify;
        }

        private void BuildRegularColoredParagraphString(List<string> rawText, Document doc)
        {
            object oCollapseEnd = Word.WdCollapseDirection.wdCollapseEnd;
            doc.Paragraphs.Add();
            Word.Range objRange = doc.Content;

            for(int i = 0; i < rawText.Count; i++)
            {
                objRange.Collapse(ref oCollapseEnd);

                if (i == 0 || i == rawText.Count - 1)
                {
                    objRange.Text = rawText[i];
                    objRange.Font.Name = "Arial";
                    objRange.Font.Size = 11;
                    objRange.Font.Color = WdColor.wdColorRed;
                    objRange.Paragraphs.Alignment = WdParagraphAlignment.wdAlignParagraphJustify;
                }
                else
                {
                    objRange.Text = rawText[i];
                    objRange.Font.Name = "Times New Roman";
                    objRange.Font.Size = 11;
                    objRange.Font.Color = WdColor.wdColorBlack;
                    objRange.Paragraphs.Alignment = WdParagraphAlignment.wdAlignParagraphJustify;
                }
            }
        }

        private void BuildBibliographyParagraphString(string rawText, Document doc)
        {
            object oCollapseEnd = Word.WdCollapseDirection.wdCollapseEnd;
            doc.Paragraphs.Add();
            Word.Range objRange = doc.Content;
            objRange.Collapse(ref oCollapseEnd);

            objRange.Text = rawText;
            objRange.Font.Name = "Verdana";
            objRange.Font.Size = 11;
            objRange.Paragraphs.Alignment = WdParagraphAlignment.wdAlignParagraphJustify;
        }

        private void BuildTableWithString(List<string> rawText, int ind , Document doc)
        {
            object oCollapseEnd = Word.WdCollapseDirection.wdCollapseEnd;
            doc.Paragraphs.Add();
            Word.Range objRange = doc.Content;
            objRange.Collapse(ref oCollapseEnd);

            // TODO : Change method to adapt to coloring

            /*objRange.Text = "[/graphic]";
            objRange.Font.Name = "Times New Roman";
            objRange.Font.Size = 9;
            objRange.Font.Color = WdColor.wdColorRose;
            objRange.InlineShapes.AddPicture("C:\\Users\\PC\\Documents\\TFG\\1578-8423-CPD-22-1-1-13image00" + ind + ".png");
            objRange.Paragraphs.Alignment = WdParagraphAlignment.wdAlignParagraphJustify;*/

            /*objRange.Collapse();
            objRange.Text = rawText + "\r[graphic href=1578-8423-CPD-22-1-1-13]";
            objRange.Font.Name = "Times New Roman";
            objRange.Font.Size = 9;
            objRange.Paragraphs.Alignment = WdParagraphAlignment.wdAlignParagraphJustify;*/

            for (int i = 0; i < rawText.Count; i++)
            {
                objRange.Collapse(ref oCollapseEnd);

                if (rawText[i].Contains("figgrp")) 
                {
                    objRange.Text = rawText[i];
                    objRange.Font.Name = "Arial";
                    objRange.Font.Size = 9;
                    objRange.Font.Color = WdColor.wdColorAqua;
                    objRange.Paragraphs.Alignment = WdParagraphAlignment.wdAlignParagraphJustify;
                }
                else if (rawText[i].Contains("label"))
                {
                    objRange.Text = rawText[i];
                    objRange.Font.Name = "Arial";
                    objRange.Font.Size = 9;
                    objRange.Font.Color = WdColor.wdColorRed;
                    objRange.Paragraphs.Alignment = WdParagraphAlignment.wdAlignParagraphJustify;
                }
                else if (rawText[i].Contains("caption"))
                {
                    objRange.Text = rawText[i];
                    objRange.Font.Name = "Arial";
                    objRange.Font.Size = 9;
                    objRange.Font.Color = WdColor.wdColorGreen;
                    objRange.Paragraphs.Alignment = WdParagraphAlignment.wdAlignParagraphJustify;
                }
                else
                {
                    objRange.Text = rawText[i];
                    objRange.Font.Name = "Times New Roman";
                    objRange.Font.Size = 9;
                    objRange.Font.Color = WdColor.wdColorBlack;
                    objRange.Paragraphs.Alignment = WdParagraphAlignment.wdAlignParagraphJustify;
                }
            }

            objRange.Collapse(ref oCollapseEnd);
            objRange.Text = "\r[graphic href=1578-8423-CPD-22-1-1-13]";
            objRange.Font.Name = "Arial";
            objRange.Font.Size = 9;
            objRange.Font.Color = WdColor.wdColorRose;
            objRange.Paragraphs.Alignment = WdParagraphAlignment.wdAlignParagraphJustify;

            //objRange.Collapse(ref oCollapseEnd);
            //objRange.Text = "\n";

            objRange.Collapse(ref oCollapseEnd);
            objRange.Text = "[/graphic]";
            objRange.Font.Name = "Arial";
            objRange.Font.Size = 9;
            objRange.Font.Color = WdColor.wdColorRose;
            objRange.InlineShapes.AddPicture("C:\\Users\\PC\\Documents\\TFG\\1578-8423-CPD-22-1-1-13image00" + ind + ".png");
            objRange.Paragraphs.Alignment = WdParagraphAlignment.wdAlignParagraphJustify;

            objRange.Collapse(ref oCollapseEnd);
            objRange.Text = "[/figgrp]";
            objRange.Font.Name = "Arial";
            objRange.Font.Size = 9;
            objRange.Font.Color = WdColor.wdColorAqua;
            objRange.Paragraphs.Alignment = WdParagraphAlignment.wdAlignParagraphJustify;
        }

        private void BuildKeywordParagraph(List<string> rawText, Document doc) 
        {
            object oCollapseEnd = Word.WdCollapseDirection.wdCollapseEnd;
            doc.Paragraphs.Add();
            Word.Range objRange = doc.Content;
            //objRange.Collapse(ref oCollapseEnd);

            for(int i = 0; i < rawText.Count; i++)
            {
                objRange.Collapse(ref oCollapseEnd);

                if (i == 0 || i == rawText.Count - 1)
                {
                    objRange.Text = rawText[i];
                    objRange.Font.Name = "Arial";
                    objRange.Font.Size = 11;
                    objRange.Font.Color = WdColor.wdColorBrown;
                    objRange.Paragraphs.Alignment = WdParagraphAlignment.wdAlignParagraphJustify;
                }
                else if (i == 1 || i == 3)
                {
                    objRange.Text = rawText[i];
                    objRange.Font.Name = "Arial";
                    objRange.Font.Size = 11;
                    objRange.Font.Color = WdColor.wdColorBlue;
                    objRange.Paragraphs.Alignment = WdParagraphAlignment.wdAlignParagraphJustify;
                }
                else if (rawText[i] == "[kwd]" || rawText[i].Contains("[/kwd]"))
                {
                    if (rawText[i] == "[kwd]" && !rawText[i-1].Contains("[/kwd]"))
                    {
                        objRange.Text = "\t" + rawText[i];
                    }
                    else
                    {
                        objRange.Text = rawText[i];
                    }
                    objRange.Font.Name = "Arial";
                    objRange.Font.Size = 11;
                    objRange.Font.Color = WdColor.wdColorGreen;
                    objRange.Paragraphs.Alignment = WdParagraphAlignment.wdAlignParagraphJustify;
                }
                else
                {
                    objRange.Text = rawText[i].Trim();
                    objRange.Font.Name = "Times New Roman";
                    objRange.Font.Size = 11;
                    objRange.Font.Color = WdColor.wdColorBlack;
                    objRange.Paragraphs.Alignment = WdParagraphAlignment.wdAlignParagraphJustify;
                }

                /*objRange.Text = rawText[i];
                objRange.Font.Name = "Arial";
                objRange.Font.Size = 11;
                objRange.Paragraphs.Alignment = WdParagraphAlignment.wdAlignParagraphJustify;*/
            }

            /*objRange.Text = rawText[i];
            objRange.Font.Name = "Ariel";
            objRange.Font.Size = 11;
            objRange.Paragraphs.Alignment = WdParagraphAlignment.wdAlignParagraphJustify;*/
        }

        private void InsertTextInNewDoc(string rawText, Document doc)
        {
            object oCollapseEnd = Word.WdCollapseDirection.wdCollapseEnd;
            doc.Paragraphs.Add();
            Word.Range objRange = doc.Content;
            objRange.Collapse(ref oCollapseEnd);

            objRange.Text = rawText;

        }

        /*private void MarkImportedDocument2()
        {
            try
            {
                bool alsoCheckRest = false;
                bool reachedBibliography = false;
                bool body = false;
                int referenceIndex = 1;
                using (var importedDoc = DocX.Load(_importedDocument))
                {
                    foreach (var sections in importedDoc.Sections)
                    {
                        foreach(var par in sections.SectionParagraphs)
                        {
                            if (par.MagicText.Any(x => x.formatting.Bold == true && x.formatting.Size > 13))
                            {
                                var markedParagraph = "[doctile]" + par.Text + "[/doctile]";
                                var replaceOptions = new StringReplaceTextOptions { SearchValue = par.Text, NewValue = markedParagraph };
                                par.ReplaceText(replaceOptions);
                                par.Alignment = Alignment.center;
                                alsoCheckRest = true;
                            }
                            if (alsoCheckRest && par.NextParagraph != null)
                            {
                                if (par.Text == "BIBLIOGRAFÍA")
                                {
                                    reachedBibliography = true;
                                }
                                if (par.MagicText.Any(x => x.formatting.Italic != true && x.formatting.Bold != true && x.formatting.Size == 12) && par.Alignment == Xceed.Document.NET.Alignment.center)
                                {
                                    string markedAuthors = MarkAuthors(par);
                                    var replaceOptions = new StringReplaceTextOptions { SearchValue = par.Text, NewValue = markedAuthors, NewFormatting = new Formatting { FontColor = Color.Black, Size = 12, Script = Script.none } };
                                    par.ReplaceText(replaceOptions);
                                    par.Alignment = Alignment.left;
                                    par.LineSpacingBefore = 6;
                                    par.LineSpacingAfter = 6;
                                }
                                else if (par.NextParagraph.Text == "INTRODUCCIÓN")
                                {
                                    string markedText = "[xmlbody]";
                                    par.InsertText(markedText);
                                    body = true;
                                }
                                else if (body && par.Text == "INTRODUCCIÓN")
                                {
                                    string markedText = "[sec sec-type=\"intro\"][sectitle]" + par.Text + "[/sectitle]";
                                    var replaceOptions = new StringReplaceTextOptions { SearchValue = par.Text, NewValue = markedText };
                                    par.ReplaceText(replaceOptions);
                                }
                                else if (body && par.Text == "MÉTODO")
                                {
                                    string markedText = "[sec sec-type=\"methods\"][sectitle]" + par.Text + "[/sectitle]";
                                    var replaceOptions = new StringReplaceTextOptions { SearchValue = par.Text, NewValue = markedText };
                                    par.ReplaceText(replaceOptions);
                                }
                                else if (par.Text.Any() && par.Text.All(x => char.IsUpper(x)) && par.MagicText.All(x => x.formatting.Bold == true && x.formatting.Size == 11))
                                {
                                    string markedResumen = "[xmlabstr language=\"es\"][sectitle]" + par.Text + "[/sectitle]";
                                    var replaceOptions = new StringReplaceTextOptions { SearchValue = par.Text, NewValue = markedResumen, NewFormatting = new Formatting { Size = 11 } };
                                    par.ReplaceText(replaceOptions);
                                    par.Alignment = Alignment.both;
                                }
                                else if (par.Alignment == Alignment.left && par.MagicText.Count > 1 && par.MagicText.Any(x => x.formatting.Bold == true && x.formatting.Size == 11))
                                {
                                    string markedKeyWords = "[kwdgrp language=\"es\"][sectitle]" + par.MagicText[0].text + "[/sectitle]";
                                    var keywords = par.MagicText[1].text.Split(';');
                                    foreach (var word in keywords)
                                    {
                                        if (keywords.Last().Equals(word))
                                        {
                                            markedKeyWords += "[kwd]" + word + "[/kwd][/kwdgrp]";
                                        }
                                        else
                                        {
                                            markedKeyWords += "[kwd]" + word + "[/kwd];";
                                        }

                                    }
                                    var replaceOptions = new StringReplaceTextOptions { SearchValue = par.Text, NewValue = markedKeyWords };
                                    par.ReplaceText(replaceOptions);                                   
                                    par.Alignment = Alignment.both;

                                }
                                else if (!reachedBibliography && par.MagicText.Any(x => x.formatting.Size >= 10 && x.formatting.Size < 12 && x.formatting.Bold != true))
                                {
                                    string markedParagraph = "[p]" + par.Text;
                                    if (par.Text[par.Text.Length - 1] == '.' && par.Text.Length > 20)
                                    {
                                        markedParagraph += "[/p]";
                                    }
                                    if (par.NextParagraph.NextParagraph.Text.Contains("Palabras clave") || par.NextParagraph.NextParagraph.Text.Contains("Key Words"))
                                    {
                                        markedParagraph += "[/xmlabstr]";
                                    }
                                    var replaceOptions = new StringReplaceTextOptions { SearchValue = par.Text, NewValue = markedParagraph };
                                    par.ReplaceText(replaceOptions);
                                }
                                else if (reachedBibliography && par.MagicText.Any(x => x.formatting.Size >= 10 && x.formatting.Size < 12 && x.formatting.Bold != true))
                                {
                                    string markedBibliographyReference = "[ref id=\"r" + referenceIndex + "\" reftype = \"journal\"][authors role = \"nd\"]";
                                    //var authors = par.MagicText[0].text.Replace("& ", string.Empty).Split(',');
                                    var authors = par.Text.Replace("& ", string.Empty).Split(',');
                                    var datePublished = authors.Select(x => x).Where(x => x.Contains('(')).FirstOrDefault();
                                    int a = Array.IndexOf(authors, datePublished);
                                    bool moreInformation = a == authors.Length;
                                    int indexParenthesisStart;
                                    int indexParenthesisEnd;

                                    if (datePublished != null)
                                    {
                                        indexParenthesisStart = datePublished.IndexOf('(');
                                        indexParenthesisEnd = datePublished.IndexOf(')');
                                        datePublished = datePublished.Substring(indexParenthesisStart + 1, (indexParenthesisEnd - indexParenthesisStart) - 1);
                                        int parenthesisIndexInAuthors = Array.IndexOf(authors, authors.FirstOrDefault(x => x.Contains("(")));
                                        for (int i = 0; i <= parenthesisIndexInAuthors; i++)
                                        {
                                            if (i == authors.Length - 1)
                                            {
                                                string author = authors[i].Split('.')[0];
                                                markedBibliographyReference += "[fname]" + author + ".[/fname][/fauthor],\t";
                                                markedBibliographyReference += "([date dateiso=\"" + (datePublished + "0000") + "\" specyear=\"" + datePublished + "\"]" + datePublished + "[/date]).";
                                                var postDateInfo = authors[i].Split('.');
                                                for (int j = 2; j < postDateInfo.Length; j++) 
                                                {
                                                    if (j == 2)
                                                    {
                                                        markedBibliographyReference += "[arttitle]" + postDateInfo[j] +"[/arttitle]";
                                                    }
                                                    else if (j == 3)
                                                    {
                                                        markedBibliographyReference += "[source]" + postDateInfo[j] + "[/source]";
                                                    }
                                                }
                                            }
                                            else if (i % 2 == 0)
                                            {
                                                markedBibliographyReference += "[pauthor][surname]" + authors[i] + "[/surname],";
                                            }
                                            else
                                            {
                                                markedBibliographyReference += "[fname]" + authors[i] + "[/fname][/fauthor],\t";
                                            }
                                        }

                                        if (moreInformation)
                                        {

                                        }
                                        referenceIndex++;

                                        var replaceOptions = new StringReplaceTextOptions { SearchValue = par.Text, NewValue = markedBibliographyReference, NewFormatting = new Formatting { Size = 11 } };
                                        par.ReplaceText(replaceOptions);
                                        par.IndentationBefore = 0;
                                        par.IndentationHanging = 0;
                                    }
                                }
                            }
                        }
                    }
                    importedDoc.SaveAs("testDocument");
                }
            } 
            catch (Exception ex)
            {
                MessageBox.Show("An error has ocurred while marking the file " + ex.ToString());
            }
        }*/

        private string MarkAuthors(Xceed.Document.NET.Paragraph par)
        {
            string[] splittedAuthors = par.Text.Split(new char[] { ',', '.', 'y' });
            List<string> authors = new List<string>();
            string author = "";
            foreach(string val in splittedAuthors) 
            {
                if (val.Trim().Length > 1 && !val.Any(char.IsDigit))
                {
                    author = "[author role=\"nd\" rid=\"aff1\" corresp=\"n\" deceased=\"n\"\r\neqcontr=\"nd\"][surname]" + val + "[/surname]";
                }
                else if (!val.Any(char.IsDigit))
                {
                    author += ", [fname]" + val +".[/fname][/author]";
                    authors.Add(author);
                }
            }

            string res = "";
            foreach(string authorFinal in authors)
            {
                res += authorFinal + Environment.NewLine;
            }
            return res;
        }

        private string MarkAuthorsInterop(string par)
        {
            string[] splittedAuthors = par.Split(new char[] { ',', '.', 'y' });
            List<string> authors = new List<string>();
            string author = "";
            int i = 0;
            foreach (string val in splittedAuthors)
            {
                if (val.Trim().Length > 1 && !val.Any(char.IsDigit))
                {
                    author = "[author role=\"nd\" rid=\"aff1\" corresp=\"n\" deceased=\"n\" eqcontr=\"nd\"][surname]" + val + "[/surname]";
                    i++;
                }
                else if (!val.Any(char.IsDigit))
                {
                    author += ", [fname]" + val + ".[/fname][/author]";
                    authors.Add(author);
                }
            }

            string res = "";
            foreach (string authorFinal in authors)
            {
                res += authorFinal + "\v";
            }
            return res;
        }

        public event PropertyChangedEventHandler? PropertyChanged;
        private void OnPropertyChanged(string propertyName)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }
}
