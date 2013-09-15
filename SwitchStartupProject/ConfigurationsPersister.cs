using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using EnvDTE;
using Newtonsoft.Json.Linq;
using Configuration = LucidConcepts.SwitchStartupProject.OptionsPage.Configuration;
using Project = LucidConcepts.SwitchStartupProject.OptionsPage.Project;

namespace LucidConcepts.SwitchStartupProject
{
    public class ConfigurationsPersister
    {
        private readonly string settingsFileExtension;
        private readonly DTE dte;
        private string settingsForSolutionFilename;
        private JObject settings;

        public ConfigurationsPersister(DTE dte, string settingsFileExtension)
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

        public IEnumerable<Configuration> GetMultiProjectConfigurations(string key)
        {
            _EnsureSettingsLoaded();
            return from configuration in settings[key].Cast<JProperty>()
                   let projects = (from project in configuration.Value
                                   select new Project(project.Value<string>()))
                   select new Configuration(configuration.Name, projects);
        }

        public void StoreMultiProjectConfigurations(string key, IList<Configuration> configurations)
        {
            _EnsureSettingsLoaded();
            settings[key] = new JObject(from c in configurations
                                        select new JProperty(c.Name, new JArray(from p in c.Projects
                                                                                select p.Name)));
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
