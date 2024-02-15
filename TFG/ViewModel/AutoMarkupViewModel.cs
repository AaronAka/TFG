using CommunityToolkit.Mvvm.Input;
using Microsoft.Win32;
using System;
using System.ComponentModel;
using System.IO;
using System.Windows;
using System.Windows.Input;
using TFG.Model;
using Application = Microsoft.Office.Interop.Word.Application;
using Document = Microsoft.Office.Interop.Word.Document;

namespace TFG.ViewModel
{
    public class AutoMarkupViewModel : INotifyPropertyChanged
    {
        private ICommand _openFileDialogCommand;
        private string _fileContent;
        private bool _enabledMarkedButton;
        private ICommand _markFileCommand;
        private string _importedDocument;
        private DocumentMarker marker;

        public AutoMarkupViewModel()
        {
            _openFileDialogCommand = new RelayCommand(ReadUserSelectedFile, ReadUserSelectedFileCanExecute);
            _markFileCommand = new RelayCommand(MarkImportedDocument, ReadUserSelectedFileCanExecute);
            _fileContent = string.Empty;
            _importedDocument = string.Empty;
            marker = new DocumentMarker();
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

                string selectedFilePath = fileDialog.FileName.ToString();

                if (!string.IsNullOrEmpty(selectedFilePath))
                {
                    Object miss = Type.Missing;
                    object readOnly = true;
                    object isVisible = false;

                    Document document;
                    Application application = new() { Visible = false };
                    MarkingConstants.FILEPATH = selectedFilePath;
                    MarkingConstants.FILENAME = Path.GetFileNameWithoutExtension(selectedFilePath);

                    document = application.Documents.Open(selectedFilePath, ref miss, ref readOnly, ref miss, ref miss, ref miss, ref miss, ref miss,
                                                            ref miss, ref miss, ref miss, ref isVisible, ref miss, ref miss, ref miss, ref miss);
                    document.ActiveWindow.Selection.WholeStory();
                    document.ActiveWindow.Selection.Copy();

                    var dataDoc = Clipboard.GetDataObject().GetData(DataFormats.Rtf).ToString();

                    if (dataDoc != null)
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
            if (MarkingConstants.FILEPATH != null)
            {
                marker = new DocumentMarker();
                marker.MarkDocument(MarkingConstants.FILEPATH);
            }
        }

        public event PropertyChangedEventHandler? PropertyChanged;

        private void OnPropertyChanged(string propertyName)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }
}