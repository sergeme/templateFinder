using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace H_IT_Tools
{
    public class Settings : INotifyPropertyChanged
    {
        public string _permMgrDefaultConfigurationFile;
        public string _permMgrOutputDirectory;
        public string _permMgrTemplateWorkBook;
            
        public event PropertyChangedEventHandler PropertyChanged;

        public string PermMgrDefaultConfigurationFile
        {
            get { return _permMgrDefaultConfigurationFile; }
            set
            {
                _permMgrDefaultConfigurationFile = value;
                OnPropertyChanged("PermMgrDefaultConfigurationFile");
            }
        }

        public string PermMgrOutputDirectory
        {
            get { return _permMgrOutputDirectory; }
            set
            {
                _permMgrOutputDirectory = value;
                OnPropertyChanged("PermMgrOutputDirectory");
            }
        }

        public string PermMgrTemplateWorkBook
        {
            get { return _permMgrTemplateWorkBook; }
            set
            {
                _permMgrTemplateWorkBook = value;
                OnPropertyChanged("PermMgrTemplateWorkBook");
            }
        }

        protected void OnPropertyChanged(string name)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
        }
    }
}
