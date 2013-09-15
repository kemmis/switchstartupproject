using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;

namespace LucidConcepts.SwitchStartupProject.OptionsPage
{
    public class Project : INotifyPropertyChanged
    {
        private string name;

        public Project(string name)
        {
            this.name = name;
        }

        public string Name
        {
            get { return name; }
            set
            {
                name = value;
                _RaisePropertyChanged("Name");
            }
        }

        protected void _RaisePropertyChanged(string propertyName)
        {
            PropertyChanged(this, new PropertyChangedEventArgs(propertyName));
        }

        public event PropertyChangedEventHandler PropertyChanged = (s, e) => { };
    }
}
