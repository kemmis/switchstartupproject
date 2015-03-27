using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using Microsoft.VisualStudio;
using Microsoft.VisualStudio.Shell.Interop;
using Newtonsoft.Json.Linq;
using Configuration = LucidConcepts.SwitchStartupProject.OptionsPage.Configuration;
using Project = LucidConcepts.SwitchStartupProject.OptionsPage.Project;

namespace LucidConcepts.SwitchStartupProject
{
    public class SettingsFileModifiedEventArgs : EventArgs
    {
        public SettingsFileModifiedEventArgs(string settingsFilename)
        {
            this.SettingsFilename = settingsFilename;
        }

        public string SettingsFilename { get; private set; }
    }

    public delegate void SettingsFileModifiedEventHandler(object sender, SettingsFileModifiedEventArgs e);

    public class ConfigurationsPersister : IVsFileChangeEvents
    {
        private const string versionKey = "Version";
        private const string singleProjectModeKey = "SingleProjectMode";
        private const string singleProjectMruCountKey = "SingleProjectMruCount";
        private const string mostRecentlyUsedListKey = "MRU";
        private const string multiProjectConfigurationsKey = "MultiProjectConfigurations";
        private const string projectsKey = "Projects";
        private const string claKey = "CommandLineArguments";
        private const string activateCommandLineArgumentsKey = "ActivateCommandLineArguments";

        private readonly string settingsFilename;
        private JObject settings;
        private readonly IVsFileChangeEx fileChangeService;
        private uint fileChangeCookie;

        public ConfigurationsPersister(string solutionFilename, string settingsFileExtension, IVsFileChangeEx fileChangeService)
        {
            this.settingsFilename = _GetSettingsFilename(solutionFilename, settingsFileExtension);
            this.fileChangeService = fileChangeService;
            _StartTrackingSettingsFile();
            Load();
        }

        public event SettingsFileModifiedEventHandler SettingsFileModified = (s, e) => { };

        public bool ConfigurationFileExists()
        {
            return File.Exists(settingsFilename);
        }

        public void Load()
        {
            settings = ConfigurationFileExists() ? JObject.Parse(File.ReadAllText(settingsFilename)) : new JObject();
        }

        public void Persist()
        {
            _StopTrackingSettingsFile();
            settings[versionKey] = 2;
            File.WriteAllText(settingsFilename, settings.ToString());
            _StartTrackingSettingsFile();
        }

        public EMode GetSingleProjectMode()
        {
            return _ExistsKey(singleProjectModeKey) ? (EMode)settings[singleProjectModeKey].Value<int>() : EMode.All;
        }

        public void StoreSingleProjectMode(EMode mode)
        {
            settings[singleProjectModeKey] = (int)mode;
        }

        public int GetSingleProjectMruCount()
        {
            return _ExistsKey(singleProjectMruCountKey) ? settings[singleProjectMruCountKey].Value<int>() : 5;
        }

        public void StoreSingleProjectMruCount(int count)
        {
            settings[singleProjectMruCountKey] = count;
        }

        public IEnumerable<string> GetSingleProjectMruList()
        {
             return _ExistsKey(mostRecentlyUsedListKey) ? settings[mostRecentlyUsedListKey].Values<string>() : Enumerable.Empty<string>();
        }

        public void StoreSingleProjectMruList(IEnumerable<string> list)
        {
            settings[mostRecentlyUsedListKey] = JArray.FromObject(list);
        }

        public IEnumerable<Configuration> GetMultiProjectConfigurations()
        {
            var version = _GetVersion();
            if (!_ExistsKey(multiProjectConfigurationsKey)) return Enumerable.Empty<Configuration>();
            if (version == 1)
            {
                return from configuration in settings[multiProjectConfigurationsKey].Cast<JProperty>()
                       let projects = (from project in configuration.Value
                                       select new Project(project.Value<string>(), ""))
                       select new Configuration(configuration.Name, projects);
            }
            return from configuration in settings[multiProjectConfigurationsKey].Cast<JProperty>()
                   let projects = (from project in configuration.Value[projectsKey].Cast<JProperty>()
                                   select new Project(project.Name, project.Value[claKey].Value<string>()))
                   select new Configuration(configuration.Name, projects);

        }

        public void StoreMultiProjectConfigurations(IList<Configuration> configurations)
        {
            settings[multiProjectConfigurationsKey] = new JObject(
                from c in configurations
                select new JProperty(c.Name, new JObject(
                    new JProperty(projectsKey, new JObject(
                        from p in c.Projects
                        select new JProperty(p.Name, new JObject(
                            new JProperty(claKey, p.CommandLineArguments))))))));
        }

        public bool GetActivateCommandLineArguments()
        {
            return _ExistsKey(activateCommandLineArgumentsKey) && settings[activateCommandLineArgumentsKey].Value<bool>();
        }

        public void StoreActivateCommandLineArguments(bool value)
        {
            settings[activateCommandLineArgumentsKey] = value;
        }

        private string _GetSettingsFilename(string solutionFilename, string settingsFileExtension)
        {
            var solutionPath = Path.GetDirectoryName(solutionFilename);
            return Path.Combine(solutionPath, Path.GetFileNameWithoutExtension(solutionFilename) + settingsFileExtension);
        }

        private int _GetVersion()
        {
            return _ExistsKey(versionKey) ? settings[versionKey].Value<int>() : 1;
        }


        private bool _ExistsKey(string key)
        {
            return settings.Children<JProperty>().Any(child => child.Name == key);
        }

        private void _StartTrackingSettingsFile()
        {
            fileChangeService.AdviseFileChange(
                settingsFilename,
                (uint)(_VSFILECHANGEFLAGS.VSFILECHG_Add | _VSFILECHANGEFLAGS.VSFILECHG_Del | _VSFILECHANGEFLAGS.VSFILECHG_Size | _VSFILECHANGEFLAGS.VSFILECHG_Time),
                this,
                out fileChangeCookie);
        }

        private void _StopTrackingSettingsFile()
        {
            fileChangeService.UnadviseFileChange(fileChangeCookie);
        }

        #region IVsFileChangeEvents Members

        public int DirectoryChanged(string pszDirectory)
        {
            return VSConstants.S_OK;
        }

        public int FilesChanged(uint cChanges, string[] rgpszFile, uint[] rggrfChange)
        {
            // Don't need to check the arguments since we ever only track the settings file
            SettingsFileModified(this, new SettingsFileModifiedEventArgs(settingsFilename));
            return VSConstants.S_OK;
        }

        #endregion
    }
}
