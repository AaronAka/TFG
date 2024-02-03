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
                bool colourString = true;

                List<string> extractedWords = new List<string>();
                foreach (Paragraph paragraph in document.Paragraphs)
                {
                    string taggedString;
                    if (paragraph.Range.Text.Length <= 4)
                    {
                        continue;
                    }

                    extractedWords.Clear();

                    if (paragraph.Range.Bold == -1 && paragraph.Range.Font.Size > 13)
                    {
                        AddDoctitleTags(extractedWords, paragraph);

                        SetStrategy(new DoctitleParagraphFormat());
                    }
                    else if (reachedBibliography)
                    {
                        extractedWords.Add("[ref id=\"r" + refIndex + "\" reftype =\"journal\"][authors role=\"nd\"]");
                        var text = paragraph.Range.Text.Replace("\r", "").Replace("\u000e", "");
                        var splitText = text.Split(',');

                        AddBibliographyTags(extractedWords, splitText);

                        //taggedString += "[/ref]";
                        extractedWords.Add("[/ref]");

                        refIndex++;
                        SetStrategy(new BibliographyParagraphFormat());
                    }
                    else if (paragraph.Range.Bold == 0 && paragraph.Range.Italic == 0 && paragraph.Range.Font.Size == 9999999 && paragraph.Alignment == WdParagraphAlignment.wdAlignParagraphCenter)
                    {
                        extractedWords = MarkAuthorsInterop(paragraph.Range.Text);
                        SetStrategy(new AuthorsParagraphFormat());
                    }
                    else if (paragraph.Range.Text == "INTRODUCCIÓN\r" || paragraph.Range.Text == "MÉTODO\r" || paragraph.Range.Text.Contains("RESULTADOS"))
                    {
                        //taggedString = "[xmlbody]\r [sec sec-type=\"intro\"][sectitle]" + paragraph.Range.Text.Replace("\r", "") + "[/sectitle]";
                        inBody = true;
                        //strings.Add("[xmlbody]\r [sec sec-type=\"intro\"]");
                        //strings.Add("[xmlbody]\r [sec sec-type=\"intro\"]");
                        AddSecTypeTags(extractedWords, paragraph);

                        colourString = paragraph.Range.Text == "MÉTODO\r" ? false : true;

                        SetStrategy(new IntroParagraphFormat());

                    }
                    else if (paragraph.Range.Case == WdCharacterCase.wdUpperCase && paragraph.Range.Bold == -1 && paragraph.Range.Font.Size == 11)
                    {
                        if (!inBody)
                        {
                            extractedWords.Add("[xmlabstr language=\"es\"][sectitle]" + paragraph.Range.Text.Replace("\r", "") + "[/sectitle]");

                            SetStrategy(new SectitleParagraphFormat());
                        }
                        else
                        {
                            extractedWords.Add("[subsec][sectitle]" + paragraph.Range.Text.Replace("\r", "") + "[/sectitle]");
                            if (extractedWords[0].Contains("BIBLIOGRAFÍA"))
                            {
                                reachedBibliography = true;
                            }

                            SetStrategy(new RegularParagraphFormat());
                        }
                    }
                    else if (paragraph.Range.Text.Length > 2 && paragraph.Range.Font.Size >= 10 && paragraph.Range.Font.Size <= 11 && paragraph.Range.Bold != 9999999)
                    {
                        var rawText = paragraph.Range.Text.Replace("\r", "").Replace("\u000e", "");
                        if (!colourString)
                        {
                            extractedWords.Add("[p]" + rawText + "[/p]");

                            SetStrategy(new RegularParagraphFormat());
                        }
                        else
                        {
                            extractedWords.Add("[p]");
                            extractedWords.Add(rawText);
                            extractedWords.Add("[/p]");

                            SetStrategy(new RegularParagraphColoured());
                        }
                    }
                    else if (paragraph.Range.Text.Contains("Tabla")  && paragraph.Range.Font.Size >= 8.5)
                    {
                        AddTableTags(tableIndex, extractedWords, paragraph);
                        SetStrategy(new TableParagraphFormat());

                        tableIndex++;

                    }
                    else if (paragraph.Range.Bold == 9999999 && paragraph.Range.Font.Size == 11 && paragraph.Range.Text.Contains(';') && paragraph.Alignment == WdParagraphAlignment.wdAlignParagraphJustify)
                    {
                        AddKeywordTags(extractedWords, paragraph);
                        SetStrategy(new KeywordParagraphFormat());

                    }
                    else if (!inBody && paragraph.Range.Bold == 0 && paragraph.Range.Font.Size == 9999999 && paragraph.Alignment == WdParagraphAlignment.wdAlignParagraphCenter)
                    {
                        AddSourceTags(extractedWords, paragraph);
                        SetStrategy(new SourcesParagraphFormat());
                    }

                    if (extractedWords.Count > 0)
                    {
                        strategy.formatParagraph(extractedWords, newDocument);
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

        private static void AddSourceTags(List<string> extractedWords, Paragraph paragraph)
        {
            var organizations = paragraph.Range.Text.Split(";");
            foreach (var organization in organizations)
            {
                var refNumber = organization.TrimStart().ElementAt(0);
                if (Char.IsDigit(refNumber))
                {
                    var organizationString = organization.Remove(0, 1);
                    extractedWords.Add("[normaff id =\"aff" + refNumber + "\" ncountry=\"Spain\" ]");
                    extractedWords.Add("[label][sup]");
                    extractedWords.Add(refNumber + "");
                    extractedWords.Add("[/sup][/label]");
                    extractedWords.Add("[orgdiv" + refNumber + "]");

                    var locationStart = organizationString.IndexOf('(');
                    var locationEnd = organizationString.IndexOf(')');
                    var location = organizationString.Substring(locationStart + 1, locationEnd - locationStart - 1).Split(',');
                    var separateByBracket = organizationString.Split('(');

                    if (separateByBracket.Length > 0 && location.Length > 1)
                    {
                        extractedWords.Add(separateByBracket[0]);
                        extractedWords.Add("[/orgdiv" + refNumber + "]");
                        extractedWords.Add("(");
                        extractedWords.Add("[city]");
                        extractedWords.Add(location[0]);
                        extractedWords.Add("[/city]");
                        extractedWords.Add(",");
                        extractedWords.Add("[country]");
                        extractedWords.Add(location[1]);
                        extractedWords.Add(").");
                        extractedWords.Add("[normaff]\r");
                    }
                }
            }
        }

        private static void AddKeywordTags(List<string> extractedWords, Paragraph paragraph)
        {
            var dividedKeywords = paragraph.Range.Text.Replace("\r", "").Split(':');
            var keywords = dividedKeywords[1].Split(';');

            extractedWords.Add("[kwdgrp language=\"es\"]");
            extractedWords.Add("[sectitle]");
            extractedWords.Add(dividedKeywords[0] + ":");
            extractedWords.Add("[/sectitle]");

            foreach (string s in keywords)
            {
                if (keywords.Last().Equals(s))
                {
                    extractedWords.Add("[kwd]");
                    extractedWords.Add(s);
                    extractedWords.Add("[/kwd]");
                    extractedWords.Add("[/kwdgrp]");
                }
                else
                {
                    extractedWords.Add("[kwd]");
                    extractedWords.Add(s);
                    extractedWords.Add("[/kwd];");
                }
            }
        }

        private static void AddTableTags(int tableIndex, List<string> extractedWords, Paragraph paragraph)
        {
            var splitString = paragraph.Range.Text.Replace("\r", "").Replace("\u000e", "").Split('.');
            //taggedString = "[label]" + splitString[0] + "[/label].";
            //taggedString += "[caption]" + splitString[1] + splitString[splitString.Length - 1] + "[/caption]";
            extractedWords.Add("[figgrp id =" + tableIndex + "]");
            extractedWords.Add("[label]");
            extractedWords.Add(splitString[0]);
            extractedWords.Add("[/label]");
            extractedWords.Add("[caption]");
            extractedWords.Add(splitString[1] + splitString[splitString.Length - 1]);
            extractedWords.Add("[/caption]");
        }

        private void AddSecTypeTags(List<string> extractedWords, Paragraph paragraph)
        {
            var secType = GetSecType(paragraph.Range.Text);
            extractedWords.Add(string.Format("[xmlbody]\r [sec sec-type=\"{0}\"]", secType));
            extractedWords.Add("[sectitle]");
            extractedWords.Add(paragraph.Range.Text.Replace("\r", ""));
            extractedWords.Add("[/sectitle]");
        }

        private string GetSecType(string text)
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

        private static void AddBibliographyTags(List<string> extractedWords, string[] splitText)
        {
            bool dateFound = false;
            bool additionalInformation = false;
            foreach (string line in splitText)
            {
                if (!dateFound && line.Contains('('))
                {
                    AddBaseBibliographyTags(extractedWords, splitText, line);
                    dateFound = true;
                    additionalInformation = !splitText.Last().Equals(line);
                }
                else if (additionalInformation)
                {
                    AddAdditionalInformation(extractedWords, line);
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
        }

        private static void AddAdditionalInformation(List<string> extractedWords, string line)
        {
            if (line.Contains('('))
            {
                // Publishing information

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
                    // Page and article id

                    //taggedString += string.Concat("[pages]", line.AsSpan(0, indexOfSeparation).Trim(), "[/pages]");
                    extractedWords.Add(string.Concat("[pages]", line.AsSpan(0, indexOfSeparation).Trim(), "[/pages]"));
                    //taggedString += string.Concat("[pubid]", line.AsSpan(indexOfSeparation + 1).Trim() , "[/pubid]");
                    extractedWords.Add(string.Concat("[pubid]", line.AsSpan(indexOfSeparation + 1).Trim(), "[/pubid]"));
                }
                else
                {
                    // Only page information

                    //taggedString += "[pages]" + line.Trim() + "[/pages]";
                    extractedWords.Add("[pages]" + line.Trim() + "[/pages]");
                }
            }
            else if (line.Contains(':'))
            {
                // doi/publication url
                //taggedString += "[pubid]" + line.Trim() + "[/pubid]";
                extractedWords.Add("[pubid]" + line.Trim() + "[/pubid]");
            }
            else
            {
                // Only basic publishing information
                //taggedString += "[volid]" + line.Trim() + "[/volid]";
                extractedWords.Add("[volid]" + line.Trim() + "[/volid]");
            }
        }

        private static void AddBaseBibliographyTags(List<string> extractedWords, string[] splitText,  string line)
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
            }
        }

        private static void AddDoctitleTags(List<string> extractedWords, Paragraph paragraph)
        {
            extractedWords.Clear();
            extractedWords.Add("[doctitle]");
            extractedWords.Add(paragraph.Range.Text.Replace("\r", "").Replace("\u000e", ""));
            extractedWords.Add("[/doctitle]");
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
