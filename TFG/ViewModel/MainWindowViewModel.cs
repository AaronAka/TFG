using CommunityToolkit.Mvvm.Input;
using Microsoft.Win32;
using Microsoft.Office.Interop.Word;
using System;
using System.Windows;
using System.Windows.Input;
using System.ComponentModel;
using Word = Microsoft.Office.Interop.Word;
using System.Linq;
using Document = Microsoft.Office.Interop.Word.Document;
using System.Collections.Generic;
using Application = Microsoft.Office.Interop.Word.Application;
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

                List<string> extractedWords = new List<string>();
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
                    extractedWords.Clear();

                    if (paragraph.Range.Bold == -1 && paragraph.Range.Font.Size > 13)
                    {
                        extractedWords.Clear();
                        extractedWords.Add("[doctitle]");
                        extractedWords.Add(paragraph.Range.Text.Replace("\r", "").Replace("\u000e", ""));
                        extractedWords.Add("[/doctitle]");

                        SetStrategy(new DoctitleParagraphFormat());
                        strategy.formatParagraph(extractedWords, newDocument);
                    }
                    else if (reachedBibliography)
                    {
                        //taggedString = "[ref id=\"r" + refIndex + "\" reftype =\"journal\"][authors role=\"nd\"]";
                        extractedWords.Add("[ref id=\"r" + refIndex + "\" reftype =\"journal\"][authors role=\"nd\"]");
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
                                        //taggedString += " [fname]" + lastName + "[/fname][/pauthor]";
                                        extractedWords.Add(" [fname]" + lastName + "[/fname][/pauthor]");
                                    }
                                    else
                                    {
                                        //taggedString += " [fname]" + lastName + "[/fname][/pauthor], ";
                                        extractedWords.Add(" [fname]" + lastName + "[/fname][/pauthor], ");
                                    }
                                    
                                    var date = decomposedLastLine[i].Replace("(", "").Replace(")", "").Trim();
                                    //taggedString += "[/authors] ";
                                    extractedWords.Add("[/authors]");
                                    //taggedString += "([date dateiso=\"" + (date + "0000") + "\" specyear=\"" + date + "\"]" + date + "[/date]).";
                                    extractedWords.Add("([date dateiso=\"" + (date + "0000") + "\" specyear=\"" + date + "\"]" + date + "[/date]).");
                                    //taggedString += " [arttitle]" + decomposedLastLine[i + 1].Trim() + ".[/arttitle]";
                                    extractedWords.Add(" [arttitle]" + decomposedLastLine[i + 1].Trim() + ".[/arttitle]");
                                    i++;

                                    if (i <= decomposedLastLine.Length - 2)
                                    {
                                        //taggedString += "[source]" + decomposedLastLine[i + 1] + "[/source],";
                                        extractedWords.Add("[source]" + decomposedLastLine[i + 1] + "[/source],");
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
                                    //taggedString += string.Concat("[volid]", line.AsSpan(0, indexOfOpenBracket).Trim(), "[/volid]");
                                    extractedWords.Add(string.Concat("[volid]", line.AsSpan(0, indexOfOpenBracket).Trim(), "[/volid]"));
                                    //taggedString += string.Concat("([issueno]", line.AsSpan(indexOfOpenBracket + 1 , indexOfEndBracket - indexOfOpenBracket - 1).Trim(), "[/issueno]),");
                                    extractedWords.Add(string.Concat("([issueno]", line.AsSpan(indexOfOpenBracket + 1, indexOfEndBracket - indexOfOpenBracket - 1).Trim(), "[/issueno]),"));
                                }
                                else if (line.Contains('–'))
                                {
                                    int indexOfSeparation = line.IndexOf('.');

                                    if (indexOfSeparation > -1)
                                    {
                                        //taggedString += string.Concat("[pages]", line.AsSpan(0, indexOfSeparation).Trim(), "[/pages]");
                                        extractedWords.Add(string.Concat("[pages]", line.AsSpan(0, indexOfSeparation).Trim(), "[/pages]"));
                                        //taggedString += string.Concat("[pubid]", line.AsSpan(indexOfSeparation + 1).Trim() , "[/pubid]");
                                        extractedWords.Add(string.Concat("[pubid]", line.AsSpan(indexOfSeparation + 1).Trim(), "[/pubid]"));
                                    } 
                                    else
                                    {
                                        //taggedString += "[pages]" + line.Trim() + "[/pages]";
                                    }
                                }
                                else if (line.Contains(':'))
                                {
                                    //taggedString += "[pubid]" + line.Trim() + "[/pubid]";
                                    extractedWords.Add("[pubid]" + line.Trim() + "[/pubid]");
                                }
                                else
                                {
                                    //taggedString += "[volid]" + line.Trim() + "[/volid]";
                                    extractedWords.Add("[volid]" + line.Trim() + "[/volid]");
                                }
                            }
                            else if (line.Length > 3 && !line.Contains('.'))
                            {
                                //taggedString += " [pauthor][surname]" + line + "[/surname],";
                                extractedWords.Add(" [pauthor][surname]" + line + "[/surname],");
                            }
                            else
                            {
                                //taggedString += " [fname]" + line + "[/fname][/pauthor],";
                                extractedWords.Add(" [fname]" + line + "[/fname][/pauthor],");
                            }
                        }

                        //taggedString += "[/ref]";
                        extractedWords.Add("[/ref]");

                        refIndex++;
                        SetStrategy(new BibliographyParagraphFormat());
                        strategy.formatParagraph(extractedWords, newDocument);
                        //BuildBibliographyParagraphString(taggedString, newDocument);
                    }
                    else if (paragraph.Range.Bold == 0 && paragraph.Range.Italic == 0 && paragraph.Range.Font.Size == 9999999 && paragraph.Alignment == WdParagraphAlignment.wdAlignParagraphCenter)
                    {
                        extractedWords = MarkAuthorsInterop(paragraph.Range.Text);
                        SetStrategy(new AuthorsParagraphFormat());
                        strategy.formatParagraph(extractedWords, newDocument);
                        //BuildAuthorsString(taggedString, newDocument);
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

                        SetStrategy(new IntroParagraphFormat());
                        strategy.formatParagraph(strings, newDocument);
                        //BuildIntroParagraphString(strings, newDocument);
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

                        SetStrategy(new IntroParagraphFormat());
                        strategy.formatParagraph(strings, newDocument);

                        //BuildIntroParagraphString(strings, newDocument);
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

                        SetStrategy(new IntroParagraphFormat());
                        strategy.formatParagraph(strings, newDocument);
                        //BuildIntroParagraphString(strings, newDocument);
                    }
                    else if (paragraph.Range.Case == WdCharacterCase.wdUpperCase && paragraph.Range.Bold == -1 && paragraph.Range.Font.Size == 11)
                    {
                        List<string> extractedText = new List<string>();
                        if (!inBody)
                        {
                            extractedText.Add("[xmlabstr language=\"es\"][sectitle]" + paragraph.Range.Text.Replace("\r", "") + "[/sectitle]");

                            SetStrategy(new SectitleParagraphFormat());
                            strategy.formatParagraph(extractedText, newDocument);
                            //BuildSectitleString(taggedString, newDocument);
                        }
                        else
                        {
                            extractedText.Add("[subsec][sectitle]" + paragraph.Range.Text.Replace("\r", "") + "[/sectitle]");
                            if (extractedText[0].Contains("BIBLIOGRAFÍA"))
                            {
                                reachedBibliography = true;
                            }
                            SetStrategy(new RegularParagraphFormat());
                            strategy.formatParagraph(extractedText, newDocument);
                            //BuildRegularParagraphString(taggedString, newDocument);
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

                        //BuildTableWithString(strings, tableIndex , newDocument);
                        SetStrategy(new TableParagraphFormat());
                        strategy.formatParagraph(strings, newDocument);
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

                        SetStrategy(new KeywordParagraphFormat());
                        strategy.formatParagraph(strings, newDocument);
                        //BuildKeywordParagraph(strings, newDocument);

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

                                SetStrategy(new SourcesParagraphFormat());
                                strategy.formatParagraph(strings, newDocument);
                                //BuildSourcesString(strings, newDocument);
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

        /*private string MarkAuthorsInterop(string par)
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
        }*/

        private List<string> MarkAuthorsInterop(string par)
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
                    authors.Add(author + "\v");
                }
            }

            /*string res = "";
            foreach (string authorFinal in authors)
            {
                res += authorFinal + "\v";
            }*/
            return authors;
        }

        public event PropertyChangedEventHandler? PropertyChanged;
        private void OnPropertyChanged(string propertyName)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }
}
