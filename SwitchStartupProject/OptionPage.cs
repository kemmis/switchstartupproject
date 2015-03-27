using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using LucidConcepts.SwitchStartupProject.OptionsPage;
using Microsoft.VisualStudio.Shell;
using Microsoft.Win32;

namespace LucidConcepts.SwitchStartupProject
{
    public class OptionPage : UIElementDialogPage
    {
        const string schemeValueName = "Scheme";
        private readonly DefaultOptionsView view;
        private bool firstLoad = true;
        private EMode mode = EMode.All;
        private int mruCount = 5;

        public OptionPage()
        {
            view = new DefaultOptionsView { DataContext = new DefaultOptionsViewModel(this) };
        }

        protected override System.Windows.UIElement Child
        {
            get { return view; }
        }

        public EMode Mode
        {
            get { return mode; }
            set {
                mode = value;
            }
        }

        public int MostRecentlyUsedCount
        {
            get { return mruCount; }
            set { 
                mruCount = value;
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

        public SwitchStartupProjectPackage.ActivityLogger Logger { get; set; }
        public void LogInfo(string message, params object[] arguments)
        {
            if (Logger == null) return;
            Logger.LogInfo(message, arguments);
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
            LogInfo("Migrating options to schema 1");
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
