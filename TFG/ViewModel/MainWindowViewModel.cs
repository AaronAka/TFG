using CommunityToolkit.Mvvm.Input;
using Microsoft.Win32;
using Microsoft.Office.Interop.Word;
using System;
using System.Windows;
using System.Windows.Input;
using System.ComponentModel;
using Word = Microsoft.Office.Interop.Word;

namespace TFG.ViewModel
{
    public class MainWindowViewModel : INotifyPropertyChanged
    {
        private ICommand _openFileDialogCommand;
        private string _fileContent;
 

        public MainWindowViewModel() 
        {
            _openFileDialogCommand = new RelayCommand(ReadUserSelectedFile, ReadUserSelectedFileCanExecute);
            _fileContent = string.Empty;
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
                OnPropertyChanged("FileContent");
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
                    Filter = "Word documents (.doc; .docx)|*.doc;*.docx"
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

                    document = application.Documents.Open(selectedFilePath, ref miss, ref readOnly, ref miss, ref miss, ref miss, ref miss, ref miss, 
                                                            ref miss, ref miss, ref miss, ref isVisible, ref miss, ref miss, ref miss, ref miss);
                    document.ActiveWindow.Selection.WholeStory();
                    document.ActiveWindow.Selection.Copy();

                    var dataDoc = "";
                    dataDoc = Clipboard.GetDataObject().GetData(DataFormats.Rtf).ToString();
                    
                    if(dataDoc != null)
                    {
                        FileContent = dataDoc;
                    }

                    application.Quit();
                }
            } 
            catch (Exception ex)
            {
                MessageBox.Show("An error has occurred while importing the file " + ex.ToString());
            }
        }

        public event PropertyChangedEventHandler? PropertyChanged;
        private void OnPropertyChanged(string propertyName)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }
}
