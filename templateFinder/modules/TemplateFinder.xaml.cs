using System;
using System.Linq;
using System.Windows;
using Forms = System.Windows.Forms;
using System.Windows.Controls;
using Word = Microsoft.Office.Interop.Word;
using System.IO;
using Microsoft.Win32;
using VB = Microsoft.VisualBasic;
using Ionic.Utils;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Threading;

namespace TemplateFinder
{
    /// <summary>
    /// Interaction logic for TemplateFinder.xaml
    /// </summary>
    public partial class WordTemplateFinder : UserControl
    {
        public string content = "Word-Templates finden";
        public string modulename = "_Word-Templates finden";
        public string textbox = "Ich bin ein String";

        public ObservableCollection<DocumentList> documents;

        public class DocumentList : INotifyPropertyChanged
        {
            private string _documentpath;
            private string _oldtemplate;
            private string _newtemplate;

            public event PropertyChangedEventHandler PropertyChanged;

            public string DocumentPath
            {
                get { return _documentpath; }

                set
                {
                    _documentpath = value;
                    OnPropertyChanged("DocumentPath");
                }
            }

            public string OldTemplate
            {
                get { return _oldtemplate; }

                set
                {
                    _oldtemplate = value;
                    OnPropertyChanged("OldTemplate");
                }
            }

            public string NewTemplate
            {
                get { return _newtemplate; }

                set
                {
                    _newtemplate = value;
                    OnPropertyChanged("NewTemplate");
                }
            }

            protected void OnPropertyChanged(string name)
            {
                PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
            }

        }

        public WordTemplateFinder()
        {
            InitializeComponent();
            documents = new ObservableCollection<DocumentList>();
            listViewFoundDocuments.ItemsSource = documents;
        }

        private void getDocumentsPath_Click(object sender, RoutedEventArgs e)
        {
            FolderBrowserDialogEx getDocumentsPath = new FolderBrowserDialogEx();
            getDocumentsPath.Description = "Zu durchsuchenden Root-Ordner wählen";
            getDocumentsPath.ShowNewFolderButton = false;
            getDocumentsPath.ShowEditBox = true;
            getDocumentsPath.ShowFullPathInEditBox = true;
            if (getDocumentsPath.ShowDialog() == Forms.DialogResult.OK)
            {
                documentsPath.Text = getDocumentsPath.SelectedPath;
            }
        }

        private void getTemplatePath_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog getTemplatePath = new OpenFileDialog();
            getTemplatePath.Filter = "Word template files (*.dot*)|*.dot;*.dotx;*.dotm";
            if (getTemplatePath.ShowDialog() == true)
            {
                newTemplatePath.Text = getTemplatePath.FileName;
            }
        }

        private void exportDataGridToExcel_Click(object sender, RoutedEventArgs e)
        {
            /*ObservableCollection<DocumentList> exportView;
            exportView = new ObservableCollection<DocumentList>();*/
            ExportToExcel<DocumentList, ObservableCollection<DocumentList>> exportToExcel = new ExportToExcel<DocumentList, ObservableCollection<DocumentList>>();
            exportToExcel.dataToPrint = documents;
            exportToExcel.GenerateReport();
        }

        private void searchFolderForDocuments_Click(object sender, RoutedEventArgs e)
        {
            if (checkTemplateForm())
            {
                searchFolderForDocuments.IsEnabled = false;
                exportDataGridToExcel.IsEnabled = false;
                documents.Clear();
                Thread subThread = new Thread(() => findTemplates(this.Dispatcher.Invoke(() => documentsPath.Text), 
                    this.Dispatcher.Invoke(() => newTemplatePath.Text), this.Dispatcher.Invoke(() => searchString.Text),
                    this.Dispatcher.Invoke(() => replaceTemplates.IsChecked.Value), this.Dispatcher.Invoke(() => hideWord.IsChecked.Value)));
                subThread.Start();
            }
            else
            {
                MessageBox.Show("Formulareingaben überprüfen");
            }
        }

        private bool checkTemplateForm()
        {
            if (string.IsNullOrWhiteSpace(documentsPath.Text))
                return false;
            else
                return true;
        }

        private void findTemplates(string searchPath, string newTemplatePath, string searchString, bool replaceTemplates, bool hideWord)
        {
            string templatelogfile = "found_templates.txt";
            string pwprotectedlogfile = "password_protected_documents.txt";
            string errordocumentslogfile = "error_documents.txt";
            string headertemplatelog = "Dokumentename;Alter Template Pfad";
            string headererrordocuments = "Dokumente die Fehler verursachten";
            string headerpwprotected = "Dokumente die Passwortgeschützt sind";
            string logrootdirectory = "C:\\_HUWIIT\\";
            string logfoldername = "logs";
            string logLine;
            string[] filepaths;

            if (!Directory.Exists(logrootdirectory + logfoldername))
            {
                Directory.CreateDirectory(logrootdirectory + logfoldername);
            }
            string[] extensions = { ".doc", ".docx" };
            filepaths = Directory.GetFiles(searchPath, "*.*", SearchOption.AllDirectories).Where(f => extensions.Contains(System.IO.Path.GetExtension(f).ToLower())).ToArray();
            this.Dispatcher.Invoke(() => StatusBar.Maximum = filepaths.Length);
            var wordApp = new Word.Application();
            wordApp.Visible = hideWord;
            
            int iteration = 0;
            foreach (string path in filepaths)
            {
                iteration++;
                Dispatcher.Invoke(() => StatusBar.Value = iteration);
                var checkprotection = MsOfficeHelper.IsProtected(path);
                if (checkprotection == true)
                {
                    logLine = path;
                    WriteLog(logLine, pwprotectedlogfile, headerpwprotected);

                }

                else if (checkprotection == false)
                {
                    string fileName = System.IO.Path.GetFileName(path);
                    if (fileName.StartsWith("~"))
                    {

                    }

                    else
                    {

                        try
                        {
                            var document = wordApp.Documents.Open(path, AddToRecentFiles: false);
                            Word.Dialog dlg = wordApp.Dialogs[Word.WdWordDialog.wdDialogDocumentStatistics];
                            var oldTemplatePath = VB.Interaction.CallByName(dlg, "Template", VB.CallType.Get, null).ToString();
                            //MessageBox.Show(searchString);
                            if (CheckTemplatePaths(oldTemplatePath, searchString) && replaceTemplates == true)
                            {
                                document.set_AttachedTemplate(newTemplatePath);
                                document.Save();
                                logLine = String.Format("{0};{1};{2}", path, oldTemplatePath, newTemplatePath);
                                WriteLog(logLine, templatelogfile, headertemplatelog);
                                Dispatcher.Invoke(() => { documents.Add(new DocumentList() { DocumentPath = path, NewTemplate = newTemplatePath, OldTemplate = oldTemplatePath }); });
                                
                            }

                            else
                            {
                                logLine = String.Format("{0};{1}", path, oldTemplatePath);
                                WriteLog(logLine, templatelogfile, headertemplatelog);
                                Dispatcher.Invoke(() => { documents.Add(new DocumentList() { DocumentPath = path, NewTemplate = "nicht ersetzt", OldTemplate = oldTemplatePath }); });
                            }
                            document.Close();
                        }

                        catch
                        {
                            logLine = String.Format("{0}", path);
                            WriteLog(logLine, errordocumentslogfile, headererrordocuments);
                        }
                    }
                }
            }
        
        wordApp.Quit();
        Dispatcher.Invoke(() => { searchFolderForDocuments.IsEnabled = true; exportDataGridToExcel.IsEnabled = true; });

        }

        private bool CheckTemplatePaths(object oldTemplate, string searchString)
        {
            string old = oldTemplate.ToString();
            bool isOnOldServer;
            old = old.ToLower();
            searchString = searchString.ToLower();
            isOnOldServer = old.StartsWith(searchString);
            return isOnOldServer;

        }

        static void WriteLog(string line, string logfile, string logheader)
        {

            string logpath = "C:\\_HUWIIT\\logs\\";
            string filepath = string.Format("{0}{1}", logpath, logfile);
            if (!File.Exists(filepath))
            {
                // Create a file to write to.
                using (StreamWriter sw = File.CreateText(filepath))
                {
                    sw.WriteLine(logheader);
                    sw.WriteLine(line);
                }
            }
            else
            {
                using (StreamWriter sw = File.AppendText(filepath))
                {
                    sw.WriteLine(line);

                }
            }
        }
    }
}
