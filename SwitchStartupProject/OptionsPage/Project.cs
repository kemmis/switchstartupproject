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
        private string commandLineArguments;

        public Project(string name, string commandLineArguments)
        {
            this.name = name;
            this.commandLineArguments = commandLineArguments;
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

        public string CommandLineArguments
        {
            get { return commandLineArguments; }
            set
            {
                commandLineArguments = value;
                _RaisePropertyChanged("CommandLineArguments");
            }
        }

        protected void _RaisePropertyChanged(string propertyName)
        {
            PropertyChanged(this, new PropertyChangedEventArgs(propertyName));
        }

        public event PropertyChangedEventHandler PropertyChanged = (s, e) => { };
    }
}
