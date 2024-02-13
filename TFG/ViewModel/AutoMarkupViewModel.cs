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
using TFG.Model;

namespace TFG.ViewModel
{
    public class AutoMarkupViewModel : INotifyPropertyChanged
    {
        private ICommand _openFileDialogCommand;
        private string _fileContent;
        private bool _enabledMarkedButton;
        private ICommand _markFileCommand;
        private string _importedDocument;
        private IParagraphFormat strategy;

        public AutoMarkupViewModel() 
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
                    Application application = new() { Visible = false };
                    _importedDocument = selectedFilePath;

                    document = application.Documents.Open(selectedFilePath, ref miss, ref readOnly, ref miss, ref miss, ref miss, ref miss, ref miss, 
                                                            ref miss, ref miss, ref miss, ref isVisible, ref miss, ref miss, ref miss, ref miss);
                    document.ActiveWindow.Selection.WholeStory();
                    document.ActiveWindow.Selection.Copy();

                    var dataDoc = Clipboard.GetDataObject().GetData(DataFormats.Rtf).ToString();
                    
                    if(dataDoc != null)
                    {
                        FileContent = dataDoc;
                        EnabledMarkButton = true;
                    }

                    application.Quit();
                }
            } 
            catch 
            {
                MessageBox.Show("An error has occurred while importing the file");
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


                var tableIndex = 1;
                var refIndex = 1;
                bool inBody = false;
                bool reachedBibliography = false;
                bool colourString = true;

                List<string> extractedWords = new List<string>();
                foreach (Paragraph paragraph in document.Paragraphs)
                {
                    string taggedString;
                    if (ParagraphTypeIdentifier.EmptyParagraph(paragraph))
                    {
                        continue;
                    }

                    extractedWords.Clear();

                    if (ParagraphTypeIdentifier.IsDoctitleParagraph(paragraph))
                    {
                        MarkDocumentFunctions.AddDoctitleTags(extractedWords, paragraph);
                        SetStrategy(new DoctitleParagraphFormat());
                    }
                    else if (reachedBibliography)
                    {
                        //extractedWords.Add("[ref id=\"r" + refIndex + "\" reftype =\"journal\"][authors role=\"nd\"]");
                        extractedWords.Add(string.Format(MarkingConstants.BIBLIOGRAPHY_REF_OPEN, refIndex));
                        var text = paragraph.Range.Text.Replace("\r", "").Replace("\u000e", "");
                        var splitText = text.Split(',');

                        MarkDocumentFunctions.AddBibliographyTags(extractedWords, splitText);

                        //taggedString += "[/ref]";
                        extractedWords.Add(MarkingConstants.BIBLIOGRAPHY_REF_CLOSE);

                        refIndex++;
                        SetStrategy(new BibliographyParagraphFormat());
                    }
                    else if (ParagraphTypeIdentifier.IsAuthorsParagraph(paragraph))
                    {
                        extractedWords = MarkDocumentFunctions.MarkAuthorsInterop(paragraph.Range.Text);
                        SetStrategy(new AuthorsParagraphFormat());
                    }
                    else if (ParagraphTypeIdentifier.IsSectitleParagraph(paragraph))
                    {
                        //taggedString = "[xmlbody]\r [sec sec-type=\"intro\"][sectitle]" + paragraph.Range.Text.Replace("\r", "") + "[/sectitle]";
                        MarkDocumentFunctions.AddSecTypeTags(extractedWords, paragraph, inBody);
                        inBody = true;

                        colourString = paragraph.Range.Text == "MÉTODO\r" ? false : true;

                        SetStrategy(new IntroParagraphFormat());

                    }
                    else if (ParagraphTypeIdentifier.IsSubsectionParagraph(paragraph))
                    {
                        if (!inBody)
                        {
                            extractedWords.Add("[xmlabstr]" + string.Format(MarkingConstants.SECTITLE, paragraph.Range.Text.Replace("\r", "")));

                            SetStrategy(new SectitleParagraphFormat());
                        }
                        else
                        {
                            extractedWords.Add("[subsec]" + string.Format(MarkingConstants.SECTITLE, paragraph.Range.Text.Replace("\r", "")));
                            if (extractedWords[0].Contains("BIBLIOGRAFÍA"))
                            {
                                reachedBibliography = true;
                            }

                            SetStrategy(new RegularParagraphFormat());
                        }
                    }
                    else if (ParagraphTypeIdentifier.IsRegularParagraph(paragraph))
                    {
                        var rawText = paragraph.Range.Text.Replace("\r", "").Replace("\u000e", "");
                        if (!colourString)
                        {
                            extractedWords.Add(string.Format(MarkingConstants.REGULAR_PARAGRAPH, rawText));

                            SetStrategy(new RegularParagraphFormat());
                        }
                        else
                        {
                            extractedWords.Add(MarkingConstants.REGULAR_PARAGRAPH_OPEN);
                            extractedWords.Add(rawText);
                            extractedWords.Add(MarkingConstants.REGULAR_PARAGRAPH_CLOSE);

                            SetStrategy(new RegularParagraphColoured());
                        }
                    }
                    else if (ParagraphTypeIdentifier.IsTableParagraph(paragraph))
                    {
                        MarkDocumentFunctions.AddTableTags(tableIndex, extractedWords, paragraph);
                        SetStrategy(new TableParagraphFormat());

                        tableIndex++;

                    }
                    else if (ParagraphTypeIdentifier.IsKeywordsParagraph(paragraph))
                    {
                        MarkDocumentFunctions.AddKeywordTags(extractedWords, paragraph);
                        SetStrategy(new KeywordParagraphFormat());

                    }
                    else if (ParagraphTypeIdentifier.IsSourcesParagraph(inBody, paragraph))
                    {
                        MarkDocumentFunctions.AddSourceTags(extractedWords, paragraph);
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

        public event PropertyChangedEventHandler? PropertyChanged;
        private void OnPropertyChanged(string propertyName)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }
}
