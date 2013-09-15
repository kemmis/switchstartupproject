using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using LucidConcepts.SwitchStartupProject.OptionsPage;
using Microsoft.VisualStudio.Shell;
using Microsoft.Win32;

namespace LucidConcepts.SwitchStartupProject
{
    public enum EOptionParameter
    {
        Mode,
        MruCount
    }

    public enum EMode
    {
        All,
        Smart,
        MostRecentlyUsed,
    }

    public class OptionsModifiedEventArgs : EventArgs
    {
        public OptionsModifiedEventArgs(EOptionParameter optionParameter)
        {
            this.OptionParameter = optionParameter;
        }

        public EOptionParameter OptionParameter { get; private set; }
    }

    public delegate void OptionsModifiedEventHandler(object sender, OptionsModifiedEventArgs e);

    public class OptionPage : UIElementDialogPage
    {
        const string schemeValueName = "Scheme";
        private readonly OptionsView view;
        private bool firstLoad = true;
        private EMode mode = EMode.All;
        private int mruCount = 5;

        public OptionPage()
        {
            view = new OptionsView { DataContext = new OptionsViewModel(this) };
        }

        protected override System.Windows.UIElement Child
        {
            get { return view; }
        }

        public event OptionsModifiedEventHandler Modified = (s, e) => { };

        public EMode Mode
        {
            get { return mode; }
            set {
                mode = value;
                Modified(this, new OptionsModifiedEventArgs(EOptionParameter.Mode));
            }
        }

        public int MostRecentlyUsedCount
        {
            get { return mruCount; }
            set { 
                mruCount = value;
                Modified(this, new OptionsModifiedEventArgs(EOptionParameter.MruCount));
            }
        }

        public override void LoadSettingsFromStorage()
        {
            if (firstLoad)
            {
                _MigrateOptions();
                firstLoad = false;
            }
            base.LoadSettingsFromStorage();
        }

        private void _MigrateOptions()
        {

            var package = (Package)GetService(typeof(Package));
            if (package != null)
            {
                using (var userRegistryRoot = package.UserRegistryRoot)
                {
                    var settingsRegistryPath = SettingsRegistryPath;
                    var registryKey = userRegistryRoot.OpenSubKey(settingsRegistryPath, true);
                    if (registryKey != null)
                    {
                        using (registryKey)
                        {
                            var scheme = (int?)registryKey.GetValue(schemeValueName) ?? 0;
                            if (scheme <= 0) _MigrateTo1(registryKey);
                        }
                    }
                }
            }
        }

        private void _MigrateTo1(RegistryKey registryKey)
        {
            const string valueMruMode = "MruMode";
            const string valueMruCount = "MruCount";
            const string valueMostRecentlyUsedCount = "MostRecentlyUsedCount";
            const string valueMode = "Mode";

            if (registryKey.GetValue(valueMruMode) != null) registryKey.DeleteValue(valueMruMode);
            var count = registryKey.GetValue(valueMruCount);
            if (count != null)
            {
                registryKey.SetValue(valueMostRecentlyUsedCount, count);
                registryKey.DeleteValue(valueMruCount);
            }
            var mode = registryKey.GetValue(valueMode);
            if (mode != null && (string)mode == "Mru")
            {
                registryKey.SetValue(valueMode, "MostRecentlyUsed");
            }
            registryKey.SetValue(schemeValueName, 1);
        }
    }
}
