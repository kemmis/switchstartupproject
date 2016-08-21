using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using Microsoft.VisualStudio;
using Microsoft.VisualStudio.Shell.Interop;
using Newtonsoft.Json.Linq;

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

        public IEnumerable<MultiProjectConfiguration> GetMultiProjectConfigurations()
        {
            var version = _GetVersion();
            if (!_ExistsKey(multiProjectConfigurationsKey)) return Enumerable.Empty<MultiProjectConfiguration>();
            if (version == 1)
            {
                return from configuration in settings[multiProjectConfigurationsKey].Cast<JProperty>()
                       let projects = (from project in configuration.Value
                                       select new MultiProjectConfigurationProject(project.Value<string>(), ""))
                       select new MultiProjectConfiguration(configuration.Name, projects.ToList());
            }
            return from configuration in settings[multiProjectConfigurationsKey].Cast<JProperty>()
                   let projects = (from project in configuration.Value[projectsKey].Cast<JProperty>()
                                   select new MultiProjectConfigurationProject(project.Name, project.Value[claKey].Value<string>()))
                   select new MultiProjectConfiguration(configuration.Name, projects.ToList());

        }

        public bool GetActivateCommandLineArguments()
        {
            return _ExistsKey(activateCommandLineArgumentsKey) && settings[activateCommandLineArgumentsKey].Value<bool>();
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
