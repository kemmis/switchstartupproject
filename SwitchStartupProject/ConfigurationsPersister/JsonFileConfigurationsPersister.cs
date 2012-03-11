using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using EnvDTE;
using Newtonsoft.Json.Linq;

namespace LucidConcepts.SwitchStartupProject.ConfigurationsPersister
{
    public class JsonFileConfigurationsPersister : IConfigurationPersister
    {
        private const string settingsFileExtension = ".startup";
        private readonly DTE dte;
        private string settingsForSolutionFilename;
        private JObject settings;

        public JsonFileConfigurationsPersister(DTE dte)
        {
            this.dte = dte;
        }

        public void Persist()
        {
            _EnsureSettingsLoaded();
            var currentSolutionFilename = dte.Solution.FullName;
            var settingsFilename = _GetSettingsFilename(currentSolutionFilename);
            File.WriteAllText(settingsFilename, settings.ToString());
        }

        public bool Exists(string key)
        {
            _EnsureSettingsLoaded();
            return settings.Children<JProperty>().Any(child => child.Name == key);
        }

        public dynamic Get(string key)
        {
            _EnsureSettingsLoaded();
            return settings[key];
        }

        public void Store(string key, dynamic value)
        {
            _EnsureSettingsLoaded();
            settings[key] = value;
        }

        public bool ExistsList(string key)
        {
            _EnsureSettingsLoaded();
            return settings.Children<JProperty>().Any(child => child.Name == key);
        }

        public dynamic GetList(string key)
        {
            _EnsureSettingsLoaded();
            return settings[key];
        }

        public void StoreList(string key, dynamic list)
        {
            _EnsureSettingsLoaded();
            settings[key] = JArray.FromObject(list);
        }

        public bool ExistsObject(string key)
        {
            _EnsureSettingsLoaded();
            return settings.Children<JProperty>().Any(child => child.Name == key);
        }

        public dynamic GetObject(string key)
        {
            _EnsureSettingsLoaded();
            return settings[key];
        }

        public void StoreObject(string key, dynamic obj)
        {
            _EnsureSettingsLoaded();
            settings[key] = JObject.FromObject(obj);
        }

        private void _EnsureSettingsLoaded()
        {
            var currentSolutionFilename = dte.Solution.FullName;
            if (settingsForSolutionFilename != currentSolutionFilename)
            {
                var settingsFilename = _GetSettingsFilename(currentSolutionFilename);
                settings = File.Exists(settingsFilename) ? JObject.Parse(File.ReadAllText(settingsFilename)) : new JObject();
                settingsForSolutionFilename = currentSolutionFilename;
            }
        }

        private string _GetSettingsFilename(string currentSolutionFilename)
        {
            var solutionPath = Path.GetDirectoryName(currentSolutionFilename);
            return Path.Combine(solutionPath, Path.GetFileNameWithoutExtension(currentSolutionFilename) + settingsFileExtension);
        }
    }
}
