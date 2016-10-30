using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows;

using EnvDTE;

using LucidConcepts.SwitchStartupProject.Helpers;

using Microsoft.VisualStudio.Shell.Interop;

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
            var newStartupProjects = dte.Solution.SolutionBuild.StartupProjects as Array;
            if (newStartupProjects == null) return;
            if (newStartupProjects.Length == 1)
            {
                var startupProject = (string)newStartupProjects.GetValue(0);
                var project = solution.Projects.Values.SingleOrDefault(p => p.Path == startupProject);
                if (project != null)
                {
                    logger.LogInfo("New startup project was activated outside of combobox: {0}", project.Name);
                    var dropDownItem = dropdownService.DropdownList.SingleOrDefault(item =>
                        (item is SingleProjectDropdownEntry) &&
                        ((item as SingleProjectDropdownEntry).Project == project)
                    ) ?? new OtherDropdownEntry(project.Name);  // if there is no dropdown entry for the single project, add an "other" entry so the name of the project is shown
                    dropdownService.CurrentDropdownValue = dropDownItem;
                    return;
                }
                logger.LogInfo("New unknown startup project was activated outside of combobox");
                dropdownService.CurrentDropdownValue = DropdownService.OtherItem;
                return;
            }

            var currentConfig = _GetCurrentlyActiveConfiguration();
            var sortedCurrentProjects = _SortedProjects(currentConfig.Projects);
            var bestMatch = (from config in solution.Configuration.MultiProjectConfigurations
                             where config.Projects.Count == currentConfig.Projects.Count
                             let score = _EqualityScore(_SortedProjects(config.Projects), sortedCurrentProjects)
                             where score >= 0.0
                             orderby score descending
                             select config)
                            .FirstOrDefault();
            if (bestMatch == null)
            {
                logger.LogInfo("New unknown multi-project startup config was activated outside of combobox");
                dropdownService.CurrentDropdownValue = DropdownService.OtherItem;
                return;
            }
            logger.LogInfo("New multi-project startup config was activated outside of combobox: {0}", bestMatch.Name);
            var dropDownValue = dropdownService.DropdownList.SingleOrDefault(item =>
                (item is MultiProjectDropdownEntry) &&
                ((item as MultiProjectDropdownEntry).Configuration == bestMatch)
            ) ?? DropdownService.OtherItem;  // if there is no dropdown entry for the single project, add an "other" entry so the name of the project is shown
            dropdownService.CurrentDropdownValue = dropDownValue;
        }

        private IEnumerable<MultiProjectConfigurationProject> _SortedProjects(IEnumerable<MultiProjectConfigurationProject> projects)
        {
            return projects.OrderBy(proj => proj.Name).ThenBy(proj => proj.CommandLineArguments);
        }

        private double _EqualityScore(IEnumerable<MultiProjectConfigurationProject> sortedAvailableProjects, IEnumerable<MultiProjectConfigurationProject> sortedCurrentProjects)
        {
            return sortedAvailableProjects.Zip(sortedCurrentProjects, (available, current) => new { Available = available, Current = current})
                .Aggregate(0.0, (accumulated, tuple) =>
                {
                    if (tuple.Available.Name != tuple.Current.Name) return double.NegativeInfinity;     // Names don't match => not equal
                    if (tuple.Available.CommandLineArguments == null) return accumulated;               // Command line arguments not configured => equality score 0
                    if (tuple.Available.CommandLineArguments != tuple.Current.CommandLineArguments)     // Command line arguments are configured and don't match => not equal
                        return double.NegativeInfinity;
                    return accumulated + 1.0;                                                           // Command line arguments match => equality score 1
                });
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

            var oldName = project.Name;
            var oldPath = project.Path;
            project.Rename();
            var newName = project.Name;
            var newPath = project.Path;

            logger.LogInfo("Renaming project {0} ({1}) into {2} ({3}) ", oldName, oldPath, newName, newPath);
            _PopulateDropdownList();
            if (solution.Configuration.MultiProjectConfigurations.Any(config => config.Projects.Any(projConfig => projConfig.Name == oldName)))
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
            if (!dropdownService.DropdownList.Contains(selectedDropdownValue))
            {
                // previously selected configuration does no longer exist
                selectedDropdownValue = null;
            }
            _ChangeStartupProject(selectedDropdownValue);
        }

        private void _LoadConfigurationAndPopulateDropdown()
        {
            solution.Configuration = solution.ConfigurationLoader.Load();
            // Check if all manually configured startup projects exist
            if (solution.Configuration.MultiProjectConfigurations.Any(config => config.Projects.Any(configProject => solution.Projects.Values.All(project => project.Name != configProject.Name))))
            {
                MessageBox.Show("The configuration file refers to inexistent projects.\nPlease check your configuration file!", "SwitchStartupProject", MessageBoxButton.OK, MessageBoxImage.Exclamation);
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
                    solution.Projects.Values.OrderBy(project => project.Name).ForEach(project => startupProjects.Add(new SingleProjectDropdownEntry(project)));
                }
                solution.Configuration.MultiProjectConfigurations.ForEach(config => startupProjects.Add(new MultiProjectDropdownEntry(config)));
            }
            dropdownService.DropdownList = startupProjects;
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

        private MultiProjectConfiguration _GetCurrentlyActiveConfiguration()
        {
            var newStartupProjects = dte.Solution.SolutionBuild.StartupProjects as Array;
            if (newStartupProjects == null) return null;
            if (solution.Projects.Count == 0) return null;

            return new MultiProjectConfiguration(null, newStartupProjects.Cast<string>().Select(projectPath =>
            {
                var project = solution.Projects.Values.SingleOrDefault(p => p.Path == projectPath);
                if (project == null) return null;
                var cla = _GetStartArgumentsOfProject(project);
                return new MultiProjectConfigurationProject(project.Name, cla);
            }).ToList());
        }

        private void _ActivateSingleProjectConfiguration(SolutionProject project)
        {
            _SuspendChangedEvent(() =>
            {
                if (project == null) return;
                dte.Solution.SolutionBuild.StartupProjects = project.Path;
            });
        }

        private void _ActivateMultiProjectConfiguration(MultiProjectConfiguration configuration)
        {
            _SuspendChangedEvent(() =>
            {
                if (configuration.Projects.Count == 1)
                {
                    // If the multi-project startup configuration contains a single project only, handle it as if it was a single-project configuration
                    var projectName = configuration.Projects.Single().Name;
                    var project = solution.Projects.Values.SingleOrDefault(p => p.Name == projectName);
                    if (project == null) return;
                    dte.Solution.SolutionBuild.StartupProjects = project.Path;
                }
                else
                {
                    // SolutionBuild.StartupProjects expects an array of objects
                    var projectPathArray = configuration.Projects.Select(projectConfig =>
                    {
                        var project = solution.Projects.Values.SingleOrDefault(p => p.Name == projectConfig.Name);
                        return project == null ? null : (object)project.Path;
                    }).ToArray();
                    dte.Solution.SolutionBuild.StartupProjects = projectPathArray;
                }

                // Set CLA
                foreach (var projectConfig in configuration.Projects)
                {
                    var project = solution.Projects.Values.SingleOrDefault(p => p.Name == projectConfig.Name);
                    if (project == null) return;
                    _SetStartArgumentsOfProject(project, projectConfig.CommandLineArguments);
                }
            });
        }

        private string _GetStartArgumentsOfProject(SolutionProject solutionProject)
        {
            var hierarchy = solutionProject.Hierarchy;
            var project = solutionProject.Project;
            var configuration = _GetActiveConfigurationOfProject(project);
            var property = _GetStartArgumentsPropertyOfConfiguration(configuration, hierarchy);
            if (property == null) return null;
            return (string)property.Value;
        }

        private void _SetStartArgumentsOfProject(SolutionProject solutionProject, string commandLineArguments)
        {
            if (commandLineArguments == null) return;
            var hierarchy = solutionProject.Hierarchy;
            var project = solutionProject.Project;
            var configurations = _GetAllConfigurationsOfProject(project);
            foreach (var configuration in configurations)
            {
                var property = _GetStartArgumentsPropertyOfConfiguration(configuration, hierarchy);
                if (property == null) continue;
                property.Value = commandLineArguments;
            }
        }

        private EnvDTE.Configuration _GetActiveConfigurationOfProject(Project project)
        {
            if (project == null) return null;
            var configurationManager = project.ConfigurationManager;
            return configurationManager == null ? null : configurationManager.ActiveConfiguration;
        }

        private IEnumerable<EnvDTE.Configuration> _GetAllConfigurationsOfProject(Project project)
        {
            if (project == null) return null;
            var configurationManager = project.ConfigurationManager;
            return configurationManager.Cast<EnvDTE.Configuration>();
        }

        private Property _GetStartArgumentsPropertyOfConfiguration(EnvDTE.Configuration configuration, IVsHierarchy projectHierarchy)
        {
            if (configuration == null || projectHierarchy == null) return null;
            var properties = configuration.Properties;
            var project = solution.Projects.GetValueOrDefault(projectHierarchy);
            if (properties == null || project == null) return null;
            return properties.Cast<Property>().FirstOrDefault(property => property.Name == project.StartArgumentsPropertyName);
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
    }
}
