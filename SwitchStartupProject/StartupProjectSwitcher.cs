using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows;

using EnvDTE;

using LucidConcepts.SwitchStartupProject.Helpers;

using Microsoft.VisualStudio.ProjectSystem.Properties;
using Microsoft.VisualStudio.ProjectSystem.Debug;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using Microsoft.VisualStudio.VCProjectEngine;

using Newtonsoft.Json;

namespace LucidConcepts.SwitchStartupProject
{
    public class StartupProjectSwitcher
    {
        private readonly DropdownService dropdownService;
        private readonly SwitchStartupProjectPackage.ActivityLogger logger;

        private Solution solution = null;
        private bool reactToChangedEvent = true;

        private readonly DTE dte;
        private readonly IVsFileChangeEx fileChangeService;

        public StartupProjectSwitcher(DropdownService dropdownService, DTE dte, IVsFileChangeEx fileChangeService, SwitchStartupProjectPackage.ActivityLogger logger)
        {
            logger.LogInfo("Entering constructor of StartupProjectSwitcher");
            this.dropdownService = dropdownService;
            dropdownService.OnListItemSelected = _ChangeStartupProject;
            dropdownService.OnConfigurationSelected = _ShowMsgOpenSolution;
            this.dte = dte;
            this.fileChangeService = fileChangeService;
            this.logger = logger;
        }

        /// <summary>
        /// Update the selection in the dropdown box when a new startup project was set
        /// </summary>
        public void UpdateStartupProject()
        {
            // When startup project is set through dropdown, don't do anything
            if (!reactToChangedEvent) return;
            // Don't react to startup project changes while opening the solution or when multi-project configurations have not yet been loaded.
            if (solution.IsOpening) return;

            // When startup project is set in solution explorer, update combobox
            var newStartupProjects = dte.Solution.SolutionBuild.StartupProjects as object[];
            if (newStartupProjects == null) return;

            var newStartupProjectPaths = newStartupProjects.Cast<string>().ToArray();
            var currentConfig = _GetCurrentlyActiveConfiguration(newStartupProjectPaths);
            var sortedCurrentProjects = _SortedProjects(currentConfig.Projects);

            // First try to find a matching multi-project dropdown entry
            var bestMatch = (from dropdownEntry in dropdownService.DropdownList
                             let multiProjectDropdownEntry = dropdownEntry as MultiProjectDropdownEntry
                             where multiProjectDropdownEntry != null
                             let startupConfig = multiProjectDropdownEntry.Configuration
                             let score = _EqualityScore(_SortedProjects(startupConfig.Projects), sortedCurrentProjects)
                             where score >= 0.0
                             orderby score descending
                             select multiProjectDropdownEntry)
                            .FirstOrDefault() as IDropdownEntry;

            // Then try to find a matching single-project dropdown entry (if feasible)
            if (bestMatch == null && newStartupProjectPaths.Length == 1)
            {
                var startupProjectPath = (string)newStartupProjectPaths.GetValue(0);
                bestMatch = (from dropdownEntry in dropdownService.DropdownList
                             let singleProjectDropdownEntry = dropdownEntry as SingleProjectDropdownEntry
                             where singleProjectDropdownEntry != null
                             where singleProjectDropdownEntry.Project.Path == startupProjectPath
                             select singleProjectDropdownEntry)
                            .FirstOrDefault();
            }
            if (bestMatch != null)
            {
                logger.LogInfo("New startup configuration was activated outside of dropdown: {0}", bestMatch.DisplayName);
                dropdownService.CurrentDropdownValue = bestMatch;
                return;
            }

            logger.LogInfo("Unknown startup configuration was activated outside of dropdown");
            dropdownService.CurrentDropdownValue = DropdownService.OtherItem;
        }

        private IEnumerable<StartupConfigurationProject> _SortedProjects(IEnumerable<StartupConfigurationProject> startupConfigurationProjects)
        {
            return startupConfigurationProjects.OrderBy(proj => proj.Project != null ? proj.Project.Path : "")
                                                .ThenBy(proj => proj.CommandLineArguments)
                                                .ThenBy(proj => proj.WorkingDirectory)
                                                .ThenBy(proj => proj.StartProject)
                                                .ThenBy(proj => proj.StartExternalProgram)
                                                .ThenBy(proj => proj.StartBrowserWithUrl)
                                                .ThenBy(proj => proj.EnableRemoteDebugging)
                                                .ThenBy(proj => proj.RemoteDebuggingMachine);
        }

        private double _EqualityScore(IEnumerable<StartupConfigurationProject> sortedAvailableProjects, IEnumerable<StartupConfigurationProject> sortedCurrentProjects)
        {
            return sortedAvailableProjects.Zip(sortedCurrentProjects, (available, current) => new { Available = available, Current = current})
                .Aggregate(0.0, (accumulated, tuple) =>
                {
                    if (tuple.Available.Project != tuple.Current.Project) return double.NegativeInfinity;   // Projects don't match => not equal
                    accumulated += _StringEqualityScore(tuple.Available, tuple.Current, proj => proj.CommandLineArguments);
                    accumulated += _StringEqualityScore(tuple.Available, tuple.Current, proj => proj.WorkingDirectory);
                    accumulated += _EqualityScore(tuple.Available, tuple.Current, proj => proj.StartProject);
                    accumulated += _StringEqualityScore(tuple.Available, tuple.Current, proj => proj.StartExternalProgram);
                    accumulated += _StringEqualityScore(tuple.Available, tuple.Current, proj => proj.StartBrowserWithUrl);
                    accumulated += _EqualityScore(tuple.Available, tuple.Current, proj => proj.EnableRemoteDebugging);
                    accumulated += _StringEqualityScore(tuple.Available, tuple.Current, proj => proj.RemoteDebuggingMachine);
                    return accumulated;
                });
        }

        private double _StringEqualityScore(StartupConfigurationProject available, StartupConfigurationProject current, Func<StartupConfigurationProject, string> getProperty)
        {
            var availableProp = getProperty(available);
            var currentProp = getProperty(current) ?? string.Empty;     // Some properties return null instead of empty string
            if (availableProp == null) return 0.0;                      // Property not configured => equality score 0
            var result = availableProp != currentProp ?
                double.NegativeInfinity :                               // Property configured but doesn't match => not equal
                1.0;                                                    // Property configured and does match => increase equality score by 1
            return result;
        }

        private double _EqualityScore<T>(StartupConfigurationProject available, StartupConfigurationProject current, Func<StartupConfigurationProject, T> getProperty)
        {
            var availableProp = getProperty(available);
            var currentProp = getProperty(current);
            if (availableProp == null) return 0.0;                      // Property not configured => equality score 0
            var result = !EqualityComparer<T>.Default.Equals(availableProp, currentProp) ?
                double.NegativeInfinity :                               // Property configured but doesn't match => not equal
                1.0;                                                    // Property configured and does match => increase equality score by 1
            return result;
        }


        // Is called before a solution and its projects are loaded.
        // Is NOT called when a new solution and project are created.
        public void BeforeOpenSolution(string solutionFileName)
        {
            logger.LogInfo("Starting to open solution: {0}", solutionFileName);
            solution = new Solution
            {
                IsOpening = true
            };
        }

        // Is called after a solution and its projects have been loaded.
        // Is also called when a new solution and project have been created.
        public void AfterOpenSolution()
        {
            solution.IsOpening = false;
            logger.LogInfo("Finished to open solution");
            if (string.IsNullOrEmpty(dte.Solution.FullName))  // This happens e.g. when creating a new website
            {
                logger.LogInfo("Solution path not yet known. Skipping initialization of configuration persister and loading of settings.");
                return;
            }
            var configurationFilename = ConfigurationLoader.GetConfigurationFilename(dte.Solution.FullName);
            var oldConfigurationFilename = ConfigurationLoader.GetOldConfigurationFilename(dte.Solution.FullName);
            solution.ConfigurationLoader = new ConfigurationLoader(configurationFilename, logger);
            solution.ConfigurationFileTracker = new ConfigurationFileTracker(configurationFilename, fileChangeService, _LoadConfigurationAndUpdateSettingsOfCurrentStartupProject);
            var configurationFileOpener = new ConfigurationFileOpener(dte, configurationFilename, oldConfigurationFilename, solution.ConfigurationLoader);
            dropdownService.OnConfigurationSelected = configurationFileOpener.Open;
            _LoadConfigurationAndPopulateDropdown();
            // Determine the currently active startup configuration and select it in the dropdown
            UpdateStartupProject();
        }

        public void BeforeCloseSolution()
        {
            logger.LogInfo("Starting to close solution");
            if (solution == null || solution.ConfigurationFileTracker == null)
            {
                return;
            }
            solution.ConfigurationFileTracker.Stop();
            solution.ConfigurationFileTracker = null;
        }

        public void AfterCloseSolution()
        {
            logger.LogInfo("Finished to close solution");
            // When solution is closed: choose no project
            dropdownService.OnConfigurationSelected = _ShowMsgOpenSolution;
            dropdownService.CurrentDropdownValue = null;
            solution.Projects.Clear();
            solution = null;
            dropdownService.DropdownList = null;
        }

        public void OpenProject(IVsHierarchy pHierarchy, bool isCreated)
        {
            // When project is opened: register it and its name

            // Filter out hierarchy elements that don't represent projects
            var project = SolutionProject.FromHierarchy(pHierarchy, dte.Solution.FullName);
            if (project == null) return;

            logger.LogInfo("{0} project: {1}", isCreated ? "Creating" : "Opening", project.Name);
            if (solution == null)   // This happens e.g. when creating a new project or website solution
            {
                solution = new Solution
                {
                    IsOpening = true,
                };
            }
            solution.Projects.Add(pHierarchy, project);
            if (!solution.IsOpening)
            {
                _PopulateDropdownList(); // when reopening a single project, refresh list
            }
        }

        public void RenameProject(IVsHierarchy pHierarchy)
        {
            var project = solution.Projects.GetValueOrDefault(pHierarchy);
            if (project == null) return;

            var showMessage = solution.Configuration.MultiProjectConfigurations.Any(config => config.Projects.Any(configProject => _ConfigRefersToProject(configProject, project)));

            var oldName = project.Name;
            var oldPath = project.Path;
            project.Rename();
            var newName = project.Name;
            var newPath = project.Path;

            logger.LogInfo("Renaming project {0} ({1}) into {2} ({3}) ", oldName, oldPath, newName, newPath);
            _PopulateDropdownList();

            if (showMessage)
            {
                MessageBox.Show("The renamed project is part of a startup configuration.\nPlease update your configuration file!", "SwitchStartupProject", MessageBoxButton.OK, MessageBoxImage.Information);
            }
        }

        public void CloseProject(IVsHierarchy pHierarchy)
        {
            var project = solution.Projects.GetValueOrDefault(pHierarchy);
            if (project == null) return;

            logger.LogInfo("Closing project: {0}", project.Name);
            // When project is closed: remove it from list of startup projects
            solution.Projects.Remove(pHierarchy);
            if ((dropdownService.CurrentDropdownValue is SingleProjectDropdownEntry) &&
                (dropdownService.CurrentDropdownValue as SingleProjectDropdownEntry).Project == project)
            {
                dropdownService.CurrentDropdownValue = null;
            }
            _PopulateDropdownList();
        }

        public void ToggleDebuggingActive(bool debuggingActive)
        {
            logger.LogInfo(debuggingActive ? "Start debugging, disable combobox" : "Stop debugging, enable combobox");
            // When debugging command UI context is activated, disable combobox, otherwise enable combobox
            dropdownService.DropdownEnabled = !debuggingActive;
        }

        private void _LoadConfigurationAndUpdateSettingsOfCurrentStartupProject()
        {
            var selectedDropdownValue = dropdownService.CurrentDropdownValue;
            _LoadConfigurationAndPopulateDropdown();
            var equivalentItem = dropdownService.DropdownList.FirstOrDefault(item => item.IsEqual(selectedDropdownValue));
            _ChangeStartupProject(equivalentItem);
        }

        private void _LoadConfigurationAndPopulateDropdown()
        {
            solution.Configuration = solution.ConfigurationLoader.Load();
            // Check if any manually configured startup project references need to be disambiguated
            var projectsByName = solution.Projects.Values.ToLookup(project => project.Name);
            var projectNamesThatNeedDisambiguation = projectsByName.Where(kvp => kvp.Count() > 1).Select(kvp => kvp.Key).ToList();
            if (projectNamesThatNeedDisambiguation.Any())
            {
                var allProjectNameReferences = (from config in solution.Configuration.MultiProjectConfigurations
                                                from proj in config.Projects
                                                select proj.NameOrPath).Distinct();
                var ambiguousProjectNameReferences = allProjectNameReferences.Intersect(projectNamesThatNeedDisambiguation).ToList();
                if (ambiguousProjectNameReferences.Any())
                {
                    var projectNameToPath = string.Join("\n", from projectName in ambiguousProjectNameReferences
                                                              from project in projectsByName[projectName]
                                                              select string.Format("• {0}", JsonConvert.ToString(project.Path)));
                    MessageBox.Show("The configuration file refers to ambiguous project names.\nPlease use either of the following project paths instead:\n\n" + projectNameToPath, "SwitchStartupProject", MessageBoxButton.OK, MessageBoxImage.Exclamation);
                }
            }
            _PopulateDropdownList();
        }

        private void _PopulateDropdownList()
        {
            var startupProjects = new List<IDropdownEntry>();
            if (solution != null)   // Solution may be null e.g. when creating a new website
            {
                if (solution.Configuration.ListAllProjects)
                {
                    var projectsByName = solution.Projects.Values.ToLookup(project => project.Name);
                    solution.Projects.Values.OrderBy(project => project.Name).ForEach(project => startupProjects.Add(new SingleProjectDropdownEntry(project)
                    {
                        Disambiguate = projectsByName[project.Name].Count() > 1
                    }));
                }
                solution.Configuration.MultiProjectConfigurations.ForEach(config => startupProjects.Add(new MultiProjectDropdownEntry(_GetStartupConfiguration(config))));
            }
            dropdownService.DropdownList = startupProjects;
        }

        private StartupConfiguration _GetStartupConfiguration(MultiProjectConfiguration config)
        {
            var startupConfigurationProjects = config.Projects.Select(configProject =>
            {
                var project = solution.Projects.Values.FirstOrDefault(solutionProject => _ConfigRefersToProject(configProject, solutionProject));
                return new StartupConfigurationProject(project, configProject.CommandLineArguments, configProject.WorkingDirectory, configProject.StartProject, configProject.StartExternalProgram, configProject.StartBrowserWithUrl, configProject.EnableRemoteDebugging, configProject.RemoteDebuggingMachine, configProject.ProfileName);
            }).ToList();
            return new StartupConfiguration(config.Name, startupConfigurationProjects);
        }

        private void _ChangeStartupProject(IDropdownEntry newStartupProject)
        {
            dropdownService.CurrentDropdownValue = newStartupProject;
            if (newStartupProject == null)
            {
                // No startup project
                _ActivateSingleProjectConfiguration(null);
                return;
            }
            if (newStartupProject is SingleProjectDropdownEntry)
            {
                // Single startup project
                var project = (newStartupProject as SingleProjectDropdownEntry).Project;
                _ActivateSingleProjectConfiguration(project);
                return;
            }
            if (newStartupProject is MultiProjectDropdownEntry)
            {
                // Multiple startup projects
                var config = (newStartupProject as MultiProjectDropdownEntry).Configuration;
                _ActivateMultiProjectConfiguration(config);
            }
        }

        private StartupConfiguration _GetCurrentlyActiveConfiguration(IEnumerable<string> newStartupProjectPaths)
        {
            if (solution.Projects.Count == 0) return null;

            return new StartupConfiguration(null, newStartupProjectPaths.Select(projectPath =>
            {
                var solutionProject = solution.Projects.Values.SingleOrDefault(p => p.Path == projectPath);
                if (solutionProject == null) return null;

                var project = solutionProject.Project;
                if (project == null) return null;

                string cla = null;
                string workingDir = null;
                bool startProject = false;
                string startExtProg = null;
                string startBrowser = null;
                bool? enableRemote = null;
                string remoteMachine = null;
                string profileName = null;

                // Handle CPS projects in a special way
                if (IsCpsProject(solutionProject.Hierarchy))
                {
                    var context = project as IVsBrowseObjectContext;
                    if (context == null)
                    {
                        // VC implements this on their DTE.Project.Object
                        context = project.Object as IVsBrowseObjectContext;
                    }
                    if (context != null)
                    {
                        var launchSettingsProvider = context.UnconfiguredProject.Services.ExportProvider.GetExportedValue<ILaunchSettingsProvider>();
                        var launchProfile = launchSettingsProvider?.CurrentSnapshot?.ActiveProfile;
                        if (launchProfile != null)
                        {
                            profileName = launchProfile.Name;
                            cla = launchProfile.CommandLineArgs;
                            workingDir = launchProfile.WorkingDirectory;
                            startProject = launchProfile.CommandName == "Project" && !launchProfile.LaunchBrowser;
                            startExtProg = launchProfile.ExecutablePath;
                            startBrowser = launchProfile.LaunchUrl;
                            //launchProfile.OtherSettings[]
                        }
                    }
                }
                // Handle VC++ projects in a special way
                else if(project.Object is VCProject vcProject)
                {
                    var vcConfiguration = vcProject.ActiveConfiguration;
                    if (vcConfiguration.DebugSettings is VCDebugSettings vcDebugSettings)
                    {
                        cla = vcDebugSettings.CommandArguments;
                        workingDir = vcDebugSettings.WorkingDirectory;

                        if (vcDebugSettings.DebuggerFlavor == eDebuggerTypes.eLocalDebugger)
                        {
                            startExtProg = vcDebugSettings.Command;
                            enableRemote = false;
                            startProject = string.IsNullOrEmpty(startExtProg);
                        }
                        else if (vcDebugSettings.DebuggerFlavor == eDebuggerTypes.eRemoteDebugger)
                        {
                            startExtProg = vcDebugSettings.RemoteCommand;
                            enableRemote = true;
                            remoteMachine = vcDebugSettings.RemoteMachine;
                            startProject = string.IsNullOrEmpty(startExtProg);
                        }
                        else if (vcDebugSettings.DebuggerFlavor == eDebuggerTypes.eWebBrowserDebugger)
                        {
                            startBrowser = vcDebugSettings.HttpUrl;
                            startExtProg = vcDebugSettings.Command;
                            enableRemote = false;
                        }
                    }
                }
                else
                {
                    var configurationManager = project.ConfigurationManager;
                    if (configurationManager == null) return null;
                    var configuration = configurationManager.ActiveConfiguration;
                    if (configuration == null) return null;
                    var properties = configuration.Properties;
                    if (properties == null) return null;


                    foreach (var property in properties.Cast<Property>())
                    {
                        if (property == null) continue;
                        if (property.Name == solutionProject.StartArgumentsPropertyName) cla = (string)property.Value;
                        if (property.Name == "StartWorkingDirectory") workingDir = (string)property.Value;
                        if (property.Name == "StartAction") startProject = 0 == (int)property.Value;
                        if (property.Name == "StartProgram") startExtProg = (string)property.Value;
                        if (property.Name == "StartURL") startBrowser = (string)property.Value;
                        if (property.Name == "RemoteDebugEnabled") enableRemote = (bool)property.Value;
                        if (property.Name == "RemoteDebugMachine") remoteMachine = (string)property.Value;
                    }
                }
                return new StartupConfigurationProject(solutionProject, cla, workingDir, startProject, startExtProg, startBrowser, enableRemote, remoteMachine, profileName);
            }).ToList());
        }

        private bool IsCpsProject(IVsHierarchy hierarchy)
        {
            if (hierarchy == null) return false;
            return hierarchy.IsCapabilityMatch("CPS");
        }

        private void _ActivateSingleProjectConfiguration(SolutionProject project)
        {
            _SuspendChangedEvent(() =>
            {
                dte.Solution.SolutionBuild.StartupProjects = project != null ? project.Path : null;
            });
        }

        private void _ActivateMultiProjectConfiguration(StartupConfiguration startupConfiguration)
        {
            var startupProjects = startupConfiguration.Projects;
            _SuspendChangedEvent(() =>
            {
                if (startupProjects.Count == 1)
                {
                    // If the multi-project startup configuration contains a single project only, handle it as if it was a single-project configuration
                    var project = startupProjects.Single().Project;
                    if (project == null)
                    {
                        MessageBox.Show("The activated configuration refers to an inexistent project.\nPlease check your configuration file!", "SwitchStartupProject", MessageBoxButton.OK, MessageBoxImage.Exclamation);
                        return;
                    }
                    dte.Solution.SolutionBuild.StartupProjects = project.Path;
                }
                else
                {
                    // SolutionBuild.StartupProjects expects an array of objects
                    var projectPathArray = (from startupProject in startupProjects
                                           let project = startupProject.Project
                                           where project != null
                                           select (object)project.Path).ToArray();
                    if (startupProjects.Count != projectPathArray.Length)
                    {
                        MessageBox.Show("The activated configuration refers to inexistent projects.\nPlease check your configuration file!", "SwitchStartupProject", MessageBoxButton.OK, MessageBoxImage.Exclamation);
                    }
                    dte.Solution.SolutionBuild.StartupProjects = projectPathArray;
                }

                // Set properties
                foreach (var startupProject in startupProjects)
                {
                    if (startupProject.Project == null) continue;
                    var solutionProject = startupProject.Project;
                    var project = solutionProject.Project;
                    if (project == null) continue;

                    // Handle CPS projects in a special way
                    if (IsCpsProject(solutionProject.Hierarchy))
                    {
                        var context = project as IVsBrowseObjectContext;
                        if (context == null)
                        {
                            // VC implements this on their DTE.Project.Object
                            context = project.Object as IVsBrowseObjectContext;
                        }
                        if (context != null)
                        {
                            var launchSettingsProvider = context.UnconfiguredProject.Services.ExportProvider.GetExportedValue<ILaunchSettingsProvider>();
                            if (launchSettingsProvider != null)
                            {
                                if (startupProject.ProfileName != null)
                                {
                                    launchSettingsProvider.SetActiveProfileAsync(startupProject.ProfileName);
                                }

                                var launchProfile = launchSettingsProvider.CurrentSnapshot?.ActiveProfile;
                                if (launchProfile != null)
                                {
                                    var writableLaunchProfile = new WritableLaunchProfile(launchProfile)
                                    {
                                        CommandLineArgs = startupProject.CommandLineArguments,
                                        WorkingDirectory = startupProject.WorkingDirectory,
                                        ExecutablePath = startupProject.StartExternalProgram,
                                        LaunchUrl = startupProject.StartBrowserWithUrl,
                                    };
                                    if (!string.IsNullOrEmpty(startupProject.StartExternalProgram))
                                    {
                                        writableLaunchProfile.CommandName = "Executable";
                                        writableLaunchProfile.LaunchBrowser = false;
                                    }
                                    if (!string.IsNullOrEmpty(startupProject.StartBrowserWithUrl))
                                    {
                                        writableLaunchProfile.LaunchBrowser = true;
                                    }
                                    if (startupProject.StartProject == true)
                                    {
                                        writableLaunchProfile.CommandName = "Project";
                                        writableLaunchProfile.LaunchBrowser = false;
                                    }

                                    launchSettingsProvider.AddOrUpdateProfileAsync(writableLaunchProfile, false);
                                }
                            }
                        }
                    }
                    // Handle VC++ projects in a special way
                    else if (project.Object is VCProject vcProject)
                    {
                        var vcConfiguration = vcProject.ActiveConfiguration;
                        if (vcConfiguration.DebugSettings is VCDebugSettings vcDebugSettings)
                        {
                            _SetPropertyValue(v => vcDebugSettings.CommandArguments = v, startupProject.CommandLineArguments);
                            _SetPropertyValue(v => vcDebugSettings.WorkingDirectory = v, startupProject.WorkingDirectory);
                            _SetPropertyValue(v =>
                            {
                                if (startupProject.EnableRemoteDebugging == true) vcDebugSettings.RemoteCommand = v;
                                else vcDebugSettings.Command = v;
                            }, startupProject.StartExternalProgram);
                            _SetPropertyValue(v => vcDebugSettings.HttpUrl = v, startupProject.StartBrowserWithUrl);
                            _SetPropertyValue(v => vcDebugSettings.RemoteMachine = v, startupProject.RemoteDebuggingMachine);

                            if (startupProject.EnableRemoteDebugging == true) vcDebugSettings.DebuggerFlavor = eDebuggerTypes.eRemoteDebugger;
                            else if (!string.IsNullOrEmpty(startupProject.StartBrowserWithUrl)) vcDebugSettings.DebuggerFlavor = eDebuggerTypes.eWebBrowserDebugger;
                            else if (startupProject.EnableRemoteDebugging == false) vcDebugSettings.DebuggerFlavor = eDebuggerTypes.eLocalDebugger;
                            else if (!string.IsNullOrEmpty(startupProject.StartExternalProgram)) vcDebugSettings.DebuggerFlavor = eDebuggerTypes.eLocalDebugger;
                            else if (startupProject.StartProject == true) vcDebugSettings.DebuggerFlavor = eDebuggerTypes.eLocalDebugger;
                        }
                    }
                    else
                    {
                        foreach (EnvDTE.Configuration configuration in project.ConfigurationManager)
                        {
                            if (configuration == null) continue;
                            var properties = configuration.Properties;
                            if (properties == null) continue;
                            foreach (var property in properties.Cast<Property>())
                            {
                                if (property == null) continue;
                                _SetPropertyValue(property, solutionProject.StartArgumentsPropertyName, startupProject.CommandLineArguments);
                                _SetPropertyValue(property, "StartWorkingDirectory", startupProject.WorkingDirectory);
                                _SetPropertyValue(property, "StartProgram", startupProject.StartExternalProgram);
                                _SetPropertyValue(property, "StartURL", startupProject.StartBrowserWithUrl);
                                _SetPropertyValue(property, "RemoteDebugEnabled", startupProject.EnableRemoteDebugging);
                                _SetPropertyValue(property, "RemoteDebugMachine", startupProject.RemoteDebuggingMachine);

                                if (!string.IsNullOrEmpty(startupProject.StartExternalProgram)) _SetPropertyValue(property, "StartAction", 1);
                                if (!string.IsNullOrEmpty(startupProject.StartBrowserWithUrl)) _SetPropertyValue(property, "StartAction", 2);
                                if (startupProject.StartProject == true) _SetPropertyValue(property, "StartAction", 0);
                            }
                        }
                    }
                }
            });
        }

        private void _SetPropertyValue(Property property, string name, object newValue)
        {
            if (property.Name != name || newValue == null) return;
            property.Value = newValue;
        }

        private void _SetPropertyValue<T>(Action<T> setter, T newValue)
        {
            if (newValue == null) return;
            setter(newValue);
        }

        private void _SuspendChangedEvent(Action action)
        {
            reactToChangedEvent = false;
            action();
            reactToChangedEvent = true;
        }

        private void _ShowMsgOpenSolution()
        {
            MessageBox.Show("Please open a solution before you configure its startup projects.\n\nIn case a solution is open, something went wrong loading it.\nMaybe it helps to delete the solution .suo file and reload the solution?", "SwitchStartupProject", MessageBoxButton.OK, MessageBoxImage.Information);
        }

        private bool _ConfigRefersToProject(MultiProjectConfigurationProject configProject, SolutionProject project)
        {
            return configProject.NameOrPath == project.Name ||
                   configProject.NameOrPath == project.Path;
        }
    }
}
