using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Input;

namespace LucidConcepts.SwitchStartupProject.OptionsPage
{
    public class OptionsViewModel : INotifyPropertyChanged
    {
        private readonly OptionPage model;
        private Configuration selectedConfiguration;
        private Project selectedProject;

        public OptionsViewModel(OptionPage model)
        {
            this.model = model;
            model.Modified += (sender, args) =>
            {
                if (args.OptionParameter == EOptionParameter.EnableMultiProjectConfiguration)
                {
                    model.LogInfo("OptionsViewModel: {0} multi project configuration UI", model.EnableMultiProjectConfiguration ? "Enabled" : "Disabled");
                    _RaisePropertyChanged("EnableMultiProjectConfiguration");
                }
            };
            LinkUrl = "https://bitbucket.org/thirteen/switchstartupproject";
            LinkCommand = new DelegateCommand(() => Process.Start(new ProcessStartInfo(new Uri(LinkUrl).AbsoluteUri)));
        }

        #region Mode

        public IEnumerable<EMode> ModeValues
        {
            get { return Enum.GetValues(typeof (EMode)).OfType<EMode>(); }
        }

        public EMode SelectedMode
        {
            get { return model.Mode; }
            set
            {
                model.Mode = value;
                _RaisePropertyChanged("SelectedMode");
                _RaisePropertyChanged("MruVisibility");
            }
        }

        public Visibility MruVisibility
        {
            get { return model.Mode == EMode.MostRecentlyUsed ? Visibility.Visible : Visibility.Hidden; }
        }

        public int MruCount
        {
            get { return model.MostRecentlyUsedCount; }
            set
            {
                model.MostRecentlyUsedCount = value;
                _RaisePropertyChanged("MruCount");
            }
        }

        #endregion

        #region Configuration

        public BindingList<Configuration> Configurations
        {
            get { return model.Configurations; }
        }

        public Configuration SelectedConfiguration
        {
            get { return selectedConfiguration; }
            set
            {
                selectedConfiguration = value;
                _RaiseConfigurationsPropertyChanged();
                selectedProject = null;
                _RaiseProjectsPropertyChanged();
            }
        }

        public bool EnableMultiProjectConfiguration
        {
            get { return model.EnableMultiProjectConfiguration; }
        }

        public bool IsConfigurationSelected
        {
            get { return SelectedConfiguration != null; }
        }

        public DelegateCommand AddConfigurationCommand { get { return new DelegateCommand(_AddConfiguration); } }
        public DelegateCommand CloneConfigurationCommand { get { return new DelegateCommand(_CloneConfiguration, _CanCloneConfiguration); } }
        public DelegateCommand DeleteConfigurationCommand { get { return new DelegateCommand(_DeleteConfiguration, _CanDeleteConfiguration); } }
        public DelegateCommand MoveConfigurationUpCommand { get { return new DelegateCommand(_MoveConfigurationUp, _CanMoveConfigurationUp); } }
        public DelegateCommand MoveConfigurationDownCommand { get { return new DelegateCommand(_MoveConfigurationDown, _CanMoveConfigurationDown); } }


        private void _RaiseConfigurationsPropertyChanged()
        {
            _RaisePropertyChanged("Configurations");
            _RaisePropertyChanged("SelectedConfiguration");
            _RaisePropertyChanged("SelectedConfigurationName");
            _RaisePropertyChanged("IsConfigurationSelected");
            CloneConfigurationCommand.RaiseCanExecuteChanged();
            DeleteConfigurationCommand.RaiseCanExecuteChanged();
            MoveConfigurationUpCommand.RaiseCanExecuteChanged();
            MoveConfigurationDownCommand.RaiseCanExecuteChanged();
            _RaisePropertyChanged("AddConfigurationCommand");
            _RaisePropertyChanged("CloneConfigurationCommand");
            _RaisePropertyChanged("DeleteConfigurationCommand");
            _RaisePropertyChanged("MoveConfigurationUpCommand");
            _RaisePropertyChanged("MoveConfigurationDownCommand");
        }

        private void _AddConfiguration()
        {
            var newConfig = new Configuration("New Configuration");
            Configurations.Add(newConfig);
            SelectedConfiguration = newConfig;
            _RaiseConfigurationsPropertyChanged();
        }

        private bool _CanCloneConfiguration()
        {
            return SelectedConfiguration != null;
        }

        private void _CloneConfiguration()
        {
            var newConfig = new Configuration(string.Format("{0} Clone", SelectedConfiguration.Name));
            SelectedConfiguration.Projects.ForEach(p => newConfig.Projects.Add(new Project(p.Name, p.CommandLineArguments)));
            Configurations.Add(newConfig);
            SelectedConfiguration = newConfig;
            _RaiseConfigurationsPropertyChanged();
        }

        private bool _CanDeleteConfiguration()
        {
            return SelectedConfiguration != null;
        }

        private void _DeleteConfiguration()
        {
            var index = Configurations.IndexOf(SelectedConfiguration);
            if (!Configurations.Remove(SelectedConfiguration)) return;
            // select item that was after the removed one, or the one before if the removed was the last item, or none if the removed one was the only item.
            var newIndex = Math.Min(index, Configurations.Count - 1);
            SelectedConfiguration = newIndex >= 0 ? Configurations[newIndex] : null;
            _RaiseConfigurationsPropertyChanged();
        }

        private bool _CanMoveConfigurationUp()
        {
            return SelectedConfiguration != null && Configurations.IndexOf(SelectedConfiguration) > 0;
        }

        private void _MoveConfigurationUp()
        {
            var configuration = SelectedConfiguration;
            var oldIndex = Configurations.IndexOf(configuration);
            Configurations.RemoveAt(oldIndex);
            Configurations.Insert(oldIndex - 1, configuration);
            SelectedConfiguration = configuration;
        }

        private bool _CanMoveConfigurationDown()
        {
            return SelectedConfiguration != null && Configurations.IndexOf(SelectedConfiguration) < Configurations.Count - 1;
        }

        private void _MoveConfigurationDown()
        {
            var configuration = selectedConfiguration;
            var oldIndex = Configurations.IndexOf(configuration);
            Configurations.RemoveAt(oldIndex);
            Configurations.Insert(oldIndex + 1, configuration);
            SelectedConfiguration = configuration;
        }

        #endregion

        #region Project

        public BindingList<Project> Projects
        {
            get
            {
                if (selectedConfiguration == null) return null;
                return selectedConfiguration.Projects;
            }
        }

        public Project SelectedProject
        {
            get { return selectedProject; }
            set
            {
                selectedProject = value;
                _RaiseProjectsPropertyChanged();
            }
        }

        public IList<string> AllProjectNames
        {
            get { return model.GetAllProjectNames(); }
        }

        public DelegateCommand AddProjectCommand { get { return new DelegateCommand(_AddProject); } }
        public DelegateCommand DeleteProjectCommand { get { return new DelegateCommand(_DeleteProject, _CanDeleteProject); } }
        public DelegateCommand MoveProjectUpCommand { get { return new DelegateCommand(_MoveProjectUp, _CanMoveProjectUp); } }
        public DelegateCommand MoveProjectDownCommand { get { return new DelegateCommand(_MoveProjectDown, _CanMoveProjectDown); } }

        private void _RaiseProjectsPropertyChanged()
        {
            _RaisePropertyChanged("Projects");
            _RaisePropertyChanged("SelectedProject");
            DeleteProjectCommand.RaiseCanExecuteChanged();
            MoveProjectUpCommand.RaiseCanExecuteChanged();
            MoveProjectDownCommand.RaiseCanExecuteChanged();
            _RaisePropertyChanged("AddProjectCommand");
            _RaisePropertyChanged("DeleteProjectCommand");
            _RaisePropertyChanged("MoveProjectUpCommand");
            _RaisePropertyChanged("MoveProjectDownCommand");
        }

        private void _AddProject()
        {
            var newProject = new Project(AllProjectNames.FirstOrDefault(), "");
            Projects.Add(newProject);
            SelectedProject = newProject;
            _RaiseProjectsPropertyChanged();
        }

        private bool _CanDeleteProject()
        {
            return SelectedProject != null;
        }
        private void _DeleteProject()
        {
            var index = Projects.IndexOf(SelectedProject);
            if (!Projects.Remove(SelectedProject)) return;
            // select item that was after the removed one, or the one before if the removed was the last item, or none if the removed one was the only item.
            var newIndex = Math.Min(index, Projects.Count - 1);
            SelectedProject = newIndex >= 0 ? Projects[newIndex] : null;
            _RaiseProjectsPropertyChanged();
        }

        private bool _CanMoveProjectUp()
        {
            return SelectedProject != null && Projects.IndexOf(SelectedProject) > 0;
        }

        private void _MoveProjectUp()
        {
            var project = SelectedProject;
            var oldIndex = Projects.IndexOf(project);
            Projects.RemoveAt(oldIndex);
            Projects.Insert(oldIndex - 1, project);
            SelectedProject = project;
        }

        private bool _CanMoveProjectDown()
        {
            return SelectedProject != null && Projects != null && Projects.IndexOf(SelectedProject) < Projects.Count - 1;
        }

        private void _MoveProjectDown()
        {
            var project = SelectedProject;
            var oldIndex = Projects.IndexOf(project);
            Projects.RemoveAt(oldIndex);
            Projects.Insert(oldIndex + 1, project);
            SelectedProject = project;
        }

        #endregion

        #region INotifyPropertyChanged

        protected void _RaisePropertyChanged(string propertyName)
        {
            PropertyChanged(this, new PropertyChangedEventArgs(propertyName));
        }

        public event PropertyChangedEventHandler PropertyChanged = (s, e) => { };

        #endregion

        #region Hyperlink

        public ICommand LinkCommand { get; private set; }
        public string LinkUrl { get; private set; }

        #endregion
    }
}
