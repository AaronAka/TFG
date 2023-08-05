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
            try
            {
                bool alsoCheckRest = false;
                bool reachedBibliography = false;
                bool body = false;
                using (var importedDoc = DocX.Load(_importedDocument))
                {
                    foreach (var fortnite in importedDoc.Sections)
                    {
                        foreach(var par in fortnite.SectionParagraphs)
                        {
                            if (par.MagicText.Any(x => x.formatting.Bold == true && x.formatting.Size > 13))
                            {
                                var markedParagraph = "[doctile]" + par.Text + "[/doctile]";
                                var replaceOptions = new StringReplaceTextOptions { SearchValue = par.Text, NewValue = markedParagraph };
                                par.ReplaceText(replaceOptions);
                                par.Alignment = Alignment.center;
                                alsoCheckRest = true;
                            }
                            if (alsoCheckRest)
                            {
                                if (par.Text.Contains("Agradecimientos") && par.MagicText.Any(x => x.formatting.Bold == true))
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
                                    var replaceOptions = new StringReplaceTextOptions { SearchValue = par.Text, NewValue = markedText };
                                    par.ReplaceText(replaceOptions);
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
                                else if (!reachedBibliography && par.MagicText.Any(x => x.formatting.Size >= 10 && x.formatting.Size < 12 && x.formatting.Bold != true && x.formatting.Italic != true))
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
                            }
                        }
                    }
                    importedDoc.SaveAs("fortnite");
                }
            } 
            catch (Exception ex)
            {
                MessageBox.Show("An error has ocurred while marking the file " + ex.ToString());
            }
        }

        private string MarkAuthors(Xceed.Document.NET.Paragraph par)
        {
            string[] splittedAuthors = par.Text.Split(new char[] { ',', '.', 'y' });
            List<string> authors = new List<string>();
            int i = 0;
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

        /* private void MarkImportedDocument()
{
   const string wordmlNamespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
   StringBuilder textBuilder = new StringBuilder();
   using (WordprocessingDocument wdDoc = WordprocessingDocument.Open(_importedDocument, false))
   {
       NameTable nt = new NameTable();
       XmlNamespaceManager nsManager = new XmlNamespaceManager(nt);
       nsManager.AddNamespace("w", wordmlNamespace);

       XmlDocument xdoc = new XmlDocument(nt);
       xdoc.Load(wdDoc.MainDocumentPart.GetStream());

       XmlNodeList paragraphNodes = xdoc.SelectNodes("//w:p", nsManager);
       var a = 2;
   }
}*/
        /*private void MarkImportedDocument()
        {
            try
            {
                using (var importedDoc = DocX.Load(_importedDocument))
                {
                    using (var newDocument = DocX.Create((@"C:\Users\PC\Desktop\pogchamp.docx")))
                    {
                        bool foundTitle = false;
                        bool foundAuthors = false;
                        var a = importedDoc.Sections;
                        foreach (var section in a)
                        {
                            foreach (var paragraph in section.SectionParagraphs)
                            {
                                if (paragraph.Text != "")
                                {
                                    if (!foundTitle)
                                    {
                                        var magicText = paragraph.MagicText;
                                        if (magicText.Count == 1 && paragraph.MagicText[0].formatting.Bold == true && magicText[0].formatting.Size > 13)
                                        {
                                            var formats = magicText[0];
                                            newDocument.InsertParagraph().Append("[doctitle]").Color(System.Drawing.Color.Purple).FontSize(formats.formatting.Size.Value);
                                            newDocument.Paragraphs[newDocument.Paragraphs.Count - 1].Append(formats.text).Color(System.Drawing.Color.Black).FontSize(formats.formatting.Size.Value);
                                            newDocument.Paragraphs[newDocument.Paragraphs.Count - 1].Append("[/doctile]").Color(System.Drawing.Color.Purple).FontSize(formats.formatting.Size.Value);
                                            newDocument.Paragraphs[newDocument.Paragraphs.Count - 1].Alignment = Xceed.Document.NET.Alignment.center;

                                            if (paragraph.NextParagraph.MagicText.Count >= 1 && paragraph.NextParagraph.MagicText[0].formatting.Bold != true)
                                            {
                                                foundTitle = true;
                                            }
                                        }
                                        else
                                        {
                                            newDocument.InsertParagraph(paragraph);
                                        }
                                    }
                                } 
                                else
                                {
                                    newDocument.InsertParagraph();
                                }
                            }
                        }

                        newDocument.Save();

                        foreach (var despacito in importedDoc.Paragraphs)
                        {

                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("An error has ocurred while marking the file " + ex.ToString());
            }
        }*/

        /*private void MarkImportedDocument()
        {
            try
            {
                using (var importedDoc = DocX.Load(_importedDocument))
                {
                    using (var newDocument = DocX.Create(@"C:\Users\PC\Desktop\pogchamp.docx"))
                    {
                        bool foundTitle = false;
                        bool foundAuthors = false;
                        foreach (var despacito in importedDoc.Paragraphs) 
                        {
                            if (!foundTitle ||!foundAuthors)
                            {
                                if (despacito.MagicText.Any())
                                {
                                    foreach (var formats in despacito.MagicText)
                                    {
                                        if (formats.formatting.Size > 13 && formats.formatting.Bold == true)
                                        {
                                            newDocument.InsertParagraph().Append("[doctitle]").Color(System.Drawing.Color.Purple).FontSize(formats.formatting.Size.Value);
                                            newDocument.Paragraphs[newDocument.Paragraphs.Count - 1].Append(formats.text).Color(System.Drawing.Color.Black).FontSize(formats.formatting.Size.Value);
                                            newDocument.Paragraphs[newDocument.Paragraphs.Count - 1].Append("[/doctile]").Color(System.Drawing.Color.Purple).FontSize(formats.formatting.Size.Value);
                                            newDocument.Paragraphs[newDocument.Paragraphs.Count - 1].Alignment = Xceed.Document.NET.Alignment.center;
                                            foundTitle = true;
                                        }
                                        else if (formats.formatting.Size == 12 && formats.formatting.Bold == null && formats.formatting.Italic == null && foundTitle)
                                        {
                                            newDocument.InsertParagraph().Append("[author role=\"nd\" rid=\"aff1\" corresp=\"n\" deceased=\"n\" eqcontr=\"nd\"][surname]");
                                            if (!formats.text.Any(char.IsDigit))
                                            {
                                                var decomposedAuthor = formats.text.Split(',').Select(x => x.Replace(" ", string.Empty)).Where(x => !string.IsNullOrEmpty(x)).ToList();
                                                if (decomposedAuthor.Count >= 2)
                                                {
                                                    newDocument.Paragraphs[newDocument.Paragraphs.Count - 1].Append(decomposedAuthor[0] + "[/surname], [fname]" + decomposedAuthor[1] + "[/fname][/author]");
                                                }
                                            }
                                        }
                                        else if (formats.formatting.Size == 11 && formats.formatting.Bold == true)
                                        {
                                            newDocument.InsertParagraph().Append("[xmlabstr language=\"es\"]").Color(System.Drawing.Color.Red);
                                            newDocument.Paragraphs[newDocument.Paragraphs.Count - 1].Append("[sectitle]").Color(System.Drawing.Color.Blue);
                                            newDocument.Paragraphs[newDocument.Paragraphs.Count - 1].Append(formats.text).Color(System.Drawing.Color.Black);
                                            newDocument.Paragraphs[newDocument.Paragraphs.Count - 1].Append("[/sectitle]").Color(System.Drawing.Color.Blue);
                                        }
                                        else if (formats.formatting.Size == 11)
                                        {
                                            newDocument.InsertParagraph().Append("[p]").Color(System.Drawing.Color.Red);
                                            newDocument.Paragraphs[newDocument.Paragraphs.Count - 1].Append(formats.text).Color(System.Drawing.Color.Black); ;
                                            newDocument.Paragraphs[newDocument.Paragraphs.Count - 1].Append("/[p]").Color(System.Drawing.Color.Red);
                                        }
                                        else
                                        {
                                            newDocument.InsertParagraph().Append(formats.text, formats.formatting);
                                        }
                                    }

                                }
                            }
                            
                            
                        }

                        newDocument.Save();
                    }
                }

                var a = 2;

            } catch (Exception ex)
            {
                MessageBox.Show("An error has ocurred while marking the file " + ex.ToString());
            }
        }*/

        public event PropertyChangedEventHandler? PropertyChanged;
        private void OnPropertyChanged(string propertyName)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }
}
