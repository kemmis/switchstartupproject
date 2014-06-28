using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using Newtonsoft.Json.Linq;
using Configuration = LucidConcepts.SwitchStartupProject.OptionsPage.Configuration;
using Project = LucidConcepts.SwitchStartupProject.OptionsPage.Project;

namespace LucidConcepts.SwitchStartupProject
{
    public class ConfigurationsPersister
    {
        private readonly string settingsFilename;
        private JObject settings;

        public ConfigurationsPersister(string solutionFilename, string settingsFileExtension)
        {
            this.settingsFilename = _GetSettingsFilename(solutionFilename, settingsFileExtension);
            Load();
        }

        public void Load()
        {
            settings = File.Exists(settingsFilename) ? JObject.Parse(File.ReadAllText(settingsFilename)) : new JObject();
        }

        public void Persist()
        {
            File.WriteAllText(settingsFilename, settings.ToString());
        }

        public bool Exists(string key)
        {
            return _ExistsKey(key);
        }

        public string Get(string key)
        {
            if (!_ExistsKey(key)) return null;
            return settings[key].Value<string>();
        }

        public void Store(string key, string value)
        {
            settings[key] = value;
        }

        public bool ExistsList(string key)
        {
            return _ExistsKey(key);
        }

        public IEnumerable<string> GetList(string key)
        {
            if (!_ExistsKey(key)) return Enumerable.Empty<string>();
            return settings[key].Values<string>();
        }

        public void StoreList(string key, IEnumerable<string> list)
        {
            settings[key] = JArray.FromObject(list);
        }

        public IEnumerable<Configuration> GetMultiProjectConfigurations(string key)
        {
            if (!_ExistsKey(key)) return Enumerable.Empty<Configuration>();
            return from configuration in settings[key].Cast<JProperty>()
                   let projects = (from project in configuration.Value
                                   select new Project(project.Value<string>()))
                   select new Configuration(configuration.Name, projects);
        }

        public void StoreMultiProjectConfigurations(string key, IList<Configuration> configurations)
        {
            settings[key] = new JObject(from c in configurations
                                        select new JProperty(c.Name, new JArray(from p in c.Projects
                                                                                select p.Name)));
        }

        private string _GetSettingsFilename(string solutionFilename, string settingsFileExtension)
        {
            var solutionPath = Path.GetDirectoryName(solutionFilename);
            return Path.Combine(solutionPath, Path.GetFileNameWithoutExtension(solutionFilename) + settingsFileExtension);
        }

        private bool _ExistsKey(string key)
        {
            return settings.Children<JProperty>().Any(child => child.Name == key);
        }

    }
}
