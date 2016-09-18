using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using Newtonsoft.Json.Linq;

namespace LucidConcepts.SwitchStartupProject
{

    public class ConfigurationLoader
    {
        private const string versionKey = "Version";
        private const string multiProjectConfigurationsKey = "MultiProjectConfigurations";
        private const string projectsKey = "Projects";
        private const string claKey = "CommandLineArguments";
        private const string activateCommandLineArgumentsKey = "ActivateCommandLineArguments";

        private readonly string configurationFilename;
        private JObject settings;

        public static string GetConfigurationFilename(string solutionFilename)
        {
            const string configurationFileExtension = ".startup.suo";
            var solutionPath = Path.GetDirectoryName(solutionFilename);
            return Path.Combine(solutionPath, Path.GetFileNameWithoutExtension(solutionFilename) + configurationFileExtension);
        }

        public ConfigurationLoader(string configurationFilename)
        {
            this.configurationFilename = configurationFilename;
            Load();
        }

        public bool ConfigurationFileExists()
        {
            return File.Exists(configurationFilename);
        }

        public void Load()
        {
            settings = ConfigurationFileExists() ? JObject.Parse(File.ReadAllText(configurationFilename)) : new JObject();
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

        private int _GetVersion()
        {
            return _ExistsKey(versionKey) ? settings[versionKey].Value<int>() : 1;
        }

        private bool _ExistsKey(string key)
        {
            return settings.Children<JProperty>().Any(child => child.Name == key);
        }
    }
}
