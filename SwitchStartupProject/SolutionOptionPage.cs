using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using LucidConcepts.SwitchStartupProject.OptionsPage;
using System.ComponentModel;

namespace LucidConcepts.SwitchStartupProject
{
    public enum EOptionParameter
    {
        Mode,
        MruCount,
        EnableSolutionConfiguration,
        MultiProjectConfigurations,
        ActivateCommandLineArguments
    }

    public enum EMode
    {
        All,
        Smart,
        MostRecentlyUsed,
        None,
    }

    public class OptionsModifiedEventArgs : EventArgs
    {
        public OptionsModifiedEventArgs(EOptionParameter optionParameter, ListChangedEventArgs listChangedEventArgs)
        {
            this.OptionParameter = optionParameter;
            this.ListChangedEventArgs = listChangedEventArgs;
        }

        public EOptionParameter OptionParameter { get; private set; }
        public ListChangedEventArgs ListChangedEventArgs { get; private set; }
    }

    public delegate void OptionsModifiedEventHandler(object sender, OptionsModifiedEventArgs e);

    public class SolutionOptionPage : UIElementDialogPage
    {
        private readonly SolutionOptionsView view;
        private EMode mode = EMode.All;
        private int mruCount = 5;
        private bool enableSolutionConfiguration = false;
        private bool activateCommandLineArguments = false;

        public SolutionOptionPage()
        {
            view = new SolutionOptionsView { DataContext = new SolutionOptionsViewModel(this) };
            Configurations = new BindingList<Configuration>();
            Configurations.ListChanged += (sender, args) => Modified(this, new OptionsModifiedEventArgs(EOptionParameter.MultiProjectConfigurations, args));
        }

        protected override System.Windows.UIElement Child
        {
            get { return view; }
        }

        public event OptionsModifiedEventHandler Modified = (s, e) => { };

        // Not being stored by automatic persistence mechanism
        [DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)]
        public EMode Mode
        {
            get { return mode; }
            set {
                mode = value;
                Modified(this, new OptionsModifiedEventArgs(EOptionParameter.Mode, null));
            }
        }

        // Not being stored by automatic persistence mechanism
        [DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)]
        public int MostRecentlyUsedCount
        {
            get { return mruCount; }
            set { 
                mruCount = value;
                Modified(this, new OptionsModifiedEventArgs(EOptionParameter.MruCount, null));
            }
        }

        // Not being stored by automatic persistence mechanism
        [DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)]
        public Func<IList<string>> GetAllProjectNames { get; set; }

        // Not being stored by automatic persistence mechanism
        [DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)]
        public BindingList<Configuration> Configurations { get; private set; }

        // Not being stored by automatic persistence mechanism
        [DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)]
        public bool ActivateCommandLineArguments
        {
            get { return activateCommandLineArguments; }
            set
            {
                activateCommandLineArguments = value;
                Modified(this, new OptionsModifiedEventArgs(EOptionParameter.ActivateCommandLineArguments, null));
            }
        }

        // Not being stored by automatic persistence mechanism
        [DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)]
        public bool EnableSolutionConfiguration
        {
            get { return enableSolutionConfiguration; }
            set
            {
                enableSolutionConfiguration = value;
                Modified(this, new OptionsModifiedEventArgs(EOptionParameter.EnableSolutionConfiguration, null));
            }
        }

        public SwitchStartupProjectPackage.ActivityLogger Logger { get; set; }
        public void LogInfo(string message, params object[] arguments)
        {
            if (Logger == null) return;
            Logger.LogInfo(message, arguments);
        }
    }
}
