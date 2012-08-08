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
        private readonly string settingsFileExtension;
        private readonly DTE dte;
        private string settingsForSolutionFilename;
        private JObject settings;

        public JsonFileConfigurationsPersister(DTE dte, string settingsFileExtension)
        {
            this.dte = dte;
            this.settingsFileExtension = settingsFileExtension;
        }

        public void Persist()
        {
            _EnsureSettingsLoaded();
            var currentSolutionFilename = dte.Solution.FullName;
            var settingsFilename = _GetSettingsFilename(currentSolutionFilename);
            File.WriteAllText(settingsFilename, settings.ToString());
        }

        public void Clean()
        {
            var currentSolutionFilename = dte.Solution.FullName;
            var settingsFilename = _GetSettingsFilename(currentSolutionFilename);
            File.Delete(settingsFilename);
        }

        public bool Exists(string key)
        {
            _EnsureSettingsLoaded();
            return settings.Children<JProperty>().Any(child => child.Name == key);
        }

        public string Get(string key)
        {
            _EnsureSettingsLoaded();
            return settings[key].Value<string>();
        }

        public void Store(string key, string value)
        {
            _EnsureSettingsLoaded();
            settings[key] = value;
        }

        public bool ExistsList(string key)
        {
            _EnsureSettingsLoaded();
            return settings.Children<JProperty>().Any(child => child.Name == key);
        }

        public IEnumerable<string> GetList(string key)
        {
            _EnsureSettingsLoaded();
            return settings[key].Values<string>();
        }

        public void StoreList(string key, IEnumerable<string> list)
        {
            _EnsureSettingsLoaded();
            settings[key] = JArray.FromObject(list);
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
