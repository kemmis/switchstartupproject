using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Windows;

using Newtonsoft.Json.Linq;

using NJsonSchema;

namespace LucidConcepts.SwitchStartupProject
{

    public class ConfigurationLoader
    {
        private const string versionKey = "Version";
        private const int knownVersion = 3;
        private const string listAllProjectsKey = "ListAllProjects";
        private const string multiProjectConfigurationsKey = "MultiProjectConfigurations";
        private const string projectsKey = "Projects";
        private const string claKey = "CommandLineArguments";
        private const string workingDirKey = "WorkingDirectory";
        private const string startProjectKey = "StartProject";
        private const string startExtProgKey = "StartExternalProgram";
        private const string startBrowserKey = "StartBrowserWithURL";
        private const string enableRemoteDebuggingKey = "EnableRemoteDebugging";
        private const string remoteDebuggingMachineKey = "RemoteDebuggingMachine";

        private readonly string configurationFilename;
        private readonly SwitchStartupProjectPackage.ActivityLogger logger;

        public static string GetConfigurationFilename(string solutionFilename)
        {
            const string configurationFileExtension = ".startup.json";
            var solutionPath = Path.GetDirectoryName(solutionFilename);
            return Path.Combine(solutionPath, solutionFilename + configurationFileExtension);
        }

        public static string GetOldConfigurationFilename(string solutionFilename)
        {
            const string oldConfigurationFileExtension = ".startup.suo";
            var solutionPath = Path.GetDirectoryName(solutionFilename);
            return Path.Combine(solutionPath, Path.GetFileNameWithoutExtension(solutionFilename) + oldConfigurationFileExtension);
        }

        public ConfigurationLoader(string configurationFilename, SwitchStartupProjectPackage.ActivityLogger logger)
        {
            this.configurationFilename = configurationFilename;
            this.logger = logger;
        }

        public Configuration Load()
        {
            logger.LogInfo("Loading configuration for solution");
            if (!_ConfigurationFileExists())
            {
                logger.LogInfo("No configuration file found, using default configuration.");
                return _GetDefaultConfiguration();
            }

            JObject settings = null;
            try
            {
                settings = JObject.Parse(File.ReadAllText(configurationFilename));
            }
            catch (Exception e)
            {
                MessageBox.Show("Could not parse configuration file.\nPlease check the syntax of your configuration file!\nUsing default configuration instead.\n\nError:\n" + e.ToString(), "SwitchStartupProject Error", MessageBoxButton.OK, MessageBoxImage.Exclamation);
                logger.LogError("Error while parsing configuration file:");
                logger.LogError(e.ToString());
                logger.LogInfo("Using default configuration.");
                return _GetDefaultConfiguration();
            }
            JsonSchema4 schema = null;
            try
            {
                var stream = Assembly.GetExecutingAssembly().GetManifestResourceStream("LucidConcepts.SwitchStartupProject.Configuration.ConfigurationSchema.json");
                using (var reader = new StreamReader(stream))
                {
                    schema = JsonSchema4.FromJson(reader.ReadToEnd());
                }
            }
            catch (Exception e)
            {
                MessageBox.Show("Could not parse schema.\nError:\n" + e.ToString(), "SwitchStartupProject Error", MessageBoxButton.OK, MessageBoxImage.Exclamation);
                logger.LogError("Error while parsing schema:");
                logger.LogError(e.ToString());
                logger.LogInfo("Using default configuration.");
                return _GetDefaultConfiguration();
            }

            var validationErrors = schema.Validate(settings);
            if (validationErrors.Any())
            {
                var messages = string.Join("\n", validationErrors.Select(err => string.Format("{0}: {1}", err.Path, err.Kind)));
                MessageBox.Show("Could not validate schema of configuration file.\nPlease check your configuration file!\nUsing default configuration instead.\n\nErrors:\n" + messages, "SwitchStartupProject Error", MessageBoxButton.OK, MessageBoxImage.Exclamation);
                logger.LogError("Could not validate schema of configuration file.");
                logger.LogInfo("Using default configuration.");
                return _GetDefaultConfiguration();
            }

            var version = _GetVersion(settings);
            if (version != knownVersion)
            {
                MessageBox.Show("Configuration file has unknown version " + version + ".\nVersion should be " + knownVersion + ".\nUsing default configuration instead.", "SwitchStartupProject Error", MessageBoxButton.OK, MessageBoxImage.Exclamation);
                logger.LogError("Unknown configuration version " + version);
                logger.LogInfo("Using default configuration.");
                return _GetDefaultConfiguration();
            }

            var listAllProjects = _GetListAllProjects(settings);
            var multiProjectConfigurations = _GetMultiProjectConfigurations(settings);
            return new Configuration(listAllProjects, multiProjectConfigurations.ToList());
        }

        public void CreateDefaultConfigurationFile()
        {
            var sb = new StringBuilder();
            sb.AppendLine("/*");
            sb.AppendLine("    This is a configuration file for the SwitchStartupProject Visual Studio Extension");
            sb.AppendLine("    See https://bitbucket.org/thirteen/switchstartupproject/src/tip/Configuration.md");
            sb.AppendLine("*/");
            sb.AppendLine("{");
            sb.AppendLine("    /*  Configuration File Version  */");
            sb.AppendLine("    \"" + versionKey + "\": 3,");
            sb.AppendLine("    ");
            sb.AppendLine("    /*  Create an item in the dropdown list for each project in the solution?  */");
            sb.AppendLine("    \"" + listAllProjectsKey + "\": true,");
            sb.AppendLine("");
            sb.AppendLine("    /*");
            sb.AppendLine("        Dictionary of named configurations with one or multiple startup projects");
            sb.AppendLine("        and optional parameters like command line arguments and working directory.");
            sb.AppendLine("        Example:");
            sb.AppendLine("");
            sb.AppendLine("        \"" + multiProjectConfigurationsKey + "\": {");
            sb.AppendLine("            \"A + B (Ext)\": {");
            sb.AppendLine("                \"" + projectsKey + "\": {");
            sb.AppendLine("                    \"MyProjectA\": {},");
            sb.AppendLine("                    \"MyProjectB\": {");
            sb.AppendLine("                        \"" + claKey + "\": \"1234\",");
            sb.AppendLine("                        \"" + workingDirKey + "\": \"%USERPROFILE%\\\\test\",");
            sb.AppendLine("                        \"" + startExtProgKey + "\": \"c:\\\\myprogram.exe\"");
            sb.AppendLine("                    }");
            sb.AppendLine("                }");
            sb.AppendLine("            },");
            sb.AppendLine("            \"A + B\": {");
            sb.AppendLine("                \"" + projectsKey + "\": {");
            sb.AppendLine("                    \"MyProjectA\": {},");
            sb.AppendLine("                    \"MyProjectB\": {");
            sb.AppendLine("                        \"" + claKey + "\": \"\",");
            sb.AppendLine("                        \"" + workingDirKey + "\": \"\",");
            sb.AppendLine("                        \"" + startProjectKey + "\": true");
            sb.AppendLine("                    }");
            sb.AppendLine("                }");
            sb.AppendLine("            },");
            sb.AppendLine("            \"D\": {");
            sb.AppendLine("                \"" + projectsKey + "\": {");
            sb.AppendLine("                    \"MyProjectD\": {}");
            sb.AppendLine("                }");
            sb.AppendLine("            }");
            sb.AppendLine("        }");
            sb.AppendLine("    */");
            sb.AppendLine("    \"" + multiProjectConfigurationsKey + "\": {}");
            sb.AppendLine("}");
            File.WriteAllText(configurationFilename, sb.ToString());
        }

        private Configuration _GetDefaultConfiguration()
        {
            return new Configuration(true, Enumerable.Empty<MultiProjectConfiguration>().ToList());
        }

        private bool _ConfigurationFileExists()
        {
            return File.Exists(configurationFilename);
        }

        private static int _GetVersion(JObject settings)
        {
            return _ExistsKey(settings, versionKey) ? settings[versionKey].Value<int>() : 0;
        }

        private static bool _GetListAllProjects(JObject settings)
        {
            return !_ExistsKey(settings, listAllProjectsKey) || settings[listAllProjectsKey].Value<bool>();
        }

        private static IEnumerable<MultiProjectConfiguration> _GetMultiProjectConfigurations(JObject settings)
        {
            if (!_ExistsKey(settings, multiProjectConfigurationsKey)) return Enumerable.Empty<MultiProjectConfiguration>();
            return from configuration in settings[multiProjectConfigurationsKey].Cast<JProperty>()
                   let projects = (from project in configuration.Value[projectsKey].Cast<JProperty>()
                                   let cla = project.Value[claKey]
                                   let workingDir = project.Value[workingDirKey]
                                   let startProject = project.Value[startProjectKey]
                                   let startExtProg = project.Value[startExtProgKey]
                                   let startBrowser = project.Value[startBrowserKey]
                                   let enableRemote = project.Value[enableRemoteDebuggingKey]
                                   let remoteMachine = project.Value[remoteDebuggingMachineKey]
                                   select new MultiProjectConfigurationProject(
                                       project.Name,
                                       cla != null ? cla.Value<string>() : null,
                                       workingDir != null ? workingDir.Value<string>() : null,
                                       startProject != null && startProject.Value<bool>(),
                                       startExtProg != null ? startExtProg.Value<string>() : null,
                                       startBrowser != null ? startBrowser.Value<string>() : null,
                                       enableRemote != null ? enableRemote.Value<bool?>() : null,
                                       remoteMachine != null ? remoteMachine.Value<string>() : null))
                   select new MultiProjectConfiguration(configuration.Name, projects.ToList());
        }

        private static bool _ExistsKey(JObject settings, string key)
        {
            return settings.Children<JProperty>().Any(child => child.Name == key);
        }
    }
}
