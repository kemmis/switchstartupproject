using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;

namespace LucidConcepts.SwitchStartupProject.OptionsPage
{
    public class Configuration : INotifyPropertyChanged
    {
        private string name;
        private readonly BindingList<Project> projects;

        public Configuration(string name)
        {
            this.name = name;
            this.projects = new BindingList<Project>();
            projects.ListChanged += (sender, args) => _RaisePropertyChanged("Projects");
        }

        public Configuration(string name, IEnumerable<Project> projects)
            : this(name)
        {
            projects.ForEach(this.projects.Add);
        }


        public string Name {
            get { return name; }
            set
            {
                name = value;
                _RaisePropertyChanged("Name");
            }
        }

        public BindingList<Project> Projects
        {
            get { return projects; }
        }

        protected void _RaisePropertyChanged(string propertyName)
        {
            PropertyChanged(this, new PropertyChangedEventArgs(propertyName));
        }

        public event PropertyChangedEventHandler PropertyChanged = (s, e) => { };
    }
}
