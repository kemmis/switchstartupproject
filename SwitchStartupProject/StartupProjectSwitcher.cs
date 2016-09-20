using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows;

using EnvDTE;

using Microsoft.VisualStudio;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;

namespace LucidConcepts.SwitchStartupProject
{
    public class StartupProjectSwitcher
    {
        private readonly DropdownService dropdownService;
        private readonly SwitchStartupProjectPackage.ActivityLogger logger;

        private readonly Dictionary<string, IVsHierarchy> name2hierarchy = new Dictionary<string, IVsHierarchy>();
        private readonly Dictionary<string, string> name2projectPath = new Dictionary<string, string>();
        private readonly Dictionary<string, string> projectPath2name = new Dictionary<string, string>();

        private ConfigurationFileTracker configurationFileTracker;
        private ConfigurationLoader configurationLoader;
        private Configuration configuration;
        private List<string> allStartupProjects;

        private bool openingSolution = false;
        private bool reactToChangedEvent = true;

        private readonly DTE dte;
        private readonly IVsFileChangeEx fileChangeService;
        private readonly ProjectHierarchyHelper projectHierarchyHelper;

        public StartupProjectSwitcher(DropdownService dropdownService, DTE dte, IVsFileChangeEx fileChangeService, ProjectHierarchyHelper project2Hierarchy, Package package, int mruCount, SwitchStartupProjectPackage.ActivityLogger logger)
        {
            logger.LogInfo("Entering constructor of StartupProjectSwitcher");
            this.dropdownService = dropdownService;
            dropdownService.OnListItemSelected = _ChangeStartupProject;
            this.dte = dte;
            this.fileChangeService = fileChangeService;
            this.projectHierarchyHelper = project2Hierarchy;
            this.logger = logger;

            allStartupProjects = new List<string>();
            configuration = null;
        }

        /// <summary>
        /// Update the selection in the dropdown box when a new startup project was set
        /// </summary>
        public void UpdateStartupProject()
        {
            // When startup project is set through dropdown, don't do anything
            if (!reactToChangedEvent) return;
            // Don't react to startup project changes while opening the solution or when multi-project configurations have not yet been loaded.
            if (openingSolution) return;

            // When startup project is set in solution explorer, update combobox
            var newStartupProjects = dte.Solution.SolutionBuild.StartupProjects as Array;
            if (newStartupProjects == null) return;
            if (newStartupProjects.Length == 1)
            {
                var startupProject = (string)newStartupProjects.GetValue(0);
                if (projectPath2name.ContainsKey(startupProject))
                {
                    var newStartupProjectName = projectPath2name[startupProject];
                    logger.LogInfo("New startup project was activated outside of combobox: {0}", newStartupProjectName);
                    dropdownService.CurrentDropdownValue = newStartupProjectName;
                    return;
                }
                logger.LogInfo("New unknown startup project was activated outside of combobox");
                dropdownService.CurrentDropdownValue = null;
                return;
            }

            var newConfig = _GetCurrentlyActiveConfiguration();
            _SelectMultiProjectConfigInDropdown(newConfig,
                success: newStartupProjectName => logger.LogInfo("New multi-project startup config was activated outside of combobox: {0}", newStartupProjectName),
                failure: () =>
                {
                    logger.LogInfo("New unknown multi-project startup config was activated outside of combobox");
                    dropdownService.CurrentDropdownValue = null;
                });
        }

        private void _SelectMultiProjectConfigInDropdown(MultiProjectConfiguration newStartupProjects, Action<string> success, Action failure)
        {
            foreach (var config in configuration.MultiProjectConfigurations.Where(config => _AreEqual(config.Projects, newStartupProjects.Projects)))
            {
                var newStartupProjectName = config.Name;
                dropdownService.CurrentDropdownValue = newStartupProjectName;
                success(newStartupProjectName);
                return; // take first match only
            }
            failure();
        }

        private bool _AreEqual(IList<MultiProjectConfigurationProject> existingConfigurationProjects, IList<MultiProjectConfigurationProject> newConfigurationProjects)
        {
            if (existingConfigurationProjects.Count != newConfigurationProjects.Count) return false;
            return existingConfigurationProjects.Zip(newConfigurationProjects,
                (existingProj, newProj) => existingProj.Name == newProj.Name &&
                                           existingProj.CommandLineArguments == newProj.CommandLineArguments
                ).All(areEqual => areEqual);
        }


        // Is called before a solution and its projects are loaded.
        // Is NOT called when a new solution and project are created.
        public void BeforeOpenSolution(string solutionFileName)
        {
            logger.LogInfo("Starting to open solution: {0}", solutionFileName);
            openingSolution = true;
        }

        // Is called after a solution and its projects have been loaded.
        // Is also called when a new solution and project have been created.
        public void AfterOpenSolution()
        {
            openingSolution = false;
            logger.LogInfo("Finished to open solution");
            if (string.IsNullOrEmpty(dte.Solution.FullName))  // This happens e.g. when creating a new website
            {
                logger.LogInfo("Solution path not yet known. Skipping initialization of configuration persister and loading of settings.");
                return;
            }
            var configurationFilename = ConfigurationLoader.GetConfigurationFilename(dte.Solution.FullName);
            configurationLoader = new ConfigurationLoader(configurationFilename, logger);
            configurationFileTracker = new ConfigurationFileTracker(configurationFilename, fileChangeService, _LoadConfigurationAndUpdateDropdown);
            _LoadConfigurationAndUpdateDropdown();
        }

        public void BeforeCloseSolution()
        {
            logger.LogInfo("Starting to close solution");
            if (configurationFileTracker == null)  // This happens e.g. when creating a new website
            {
                return;
            }
            configurationFileTracker.Stop();
            configurationFileTracker = null;
        }

        public void AfterCloseSolution()
        {
            logger.LogInfo("Finished to close solution");
            // When solution is closed: choose no project
            dropdownService.CurrentDropdownValue = null;
            configurationLoader = null;
            _ClearProjects();
        }

        public void OpenProject(IVsHierarchy pHierarchy)
        {
            // When project is opened: register it and its name
            var name = _GetProjectStringProperty(pHierarchy, __VSHPROPID.VSHPROPID_Name);
            var typeName = _GetProjectStringProperty(pHierarchy, __VSHPROPID.VSHPROPID_TypeName);
            var caption = _GetProjectStringProperty(pHierarchy, __VSHPROPID.VSHPROPID_Caption);
            var guid = _GetProjectGuidProperty(pHierarchy, __VSHPROPID.VSHPROPID_TypeGuid);

            // Filter out hierarchy elements that don't represent projects
            if (name == null || typeName == null || caption == null || guid == null) return;
            
            logger.LogInfo("Opening project: {0}", name);

            var projectTypeGuids = _GetProjectTypeGuids(pHierarchy);

            var isWebSiteProject = projectTypeGuids.Contains(GuidList.guidWebSite);

            _AddProject(name, pHierarchy, isWebSiteProject);
            allStartupProjects.Add(name);

            if (!openingSolution)
            {
                _PopulateDropdownList(); // when reopening a single project, refresh list
            }
        }

        public void RenameProject(IVsHierarchy pHierarchy)
        {
            var newName = _GetProjectName(pHierarchy);
            if (newName == null) return;
            if (!name2hierarchy.ContainsValue(pHierarchy)) return;
            var oldName = name2hierarchy.Single(kvp => kvp.Value == pHierarchy).Key;
            var oldPath = name2projectPath[oldName];
            var projectTypeGuids = _GetProjectTypeGuids(pHierarchy);
            var isWebSiteProject = projectTypeGuids.Contains(GuidList.guidWebSite);
            // Website projects need to be set using the full path
            var newPath = isWebSiteProject ? _GetAbsolutePath(pHierarchy) : _GetPathRelativeToSolution(pHierarchy);

            logger.LogInfo("Renaming project {0} ({1}) into {2} ({3}) ", oldName, oldPath, newName, newPath);
            var reselectRenamedProject = dropdownService.CurrentDropdownValue == oldName;
            _RenameProject(pHierarchy, oldName, oldPath, newName, newPath);
            if (reselectRenamedProject)
            {
                dropdownService.CurrentDropdownValue = newName;
            }
        }

        public void CloseProject(IVsHierarchy pHierarchy, string projectName)
        {
            logger.LogInfo("Closing project: {0}", projectName);
            // When project is closed: remove it from list of startup projects (if it was in there)
            if (allStartupProjects.Contains(projectName))
            {
                if (dropdownService.CurrentDropdownValue == projectName)
                {
                    dropdownService.CurrentDropdownValue = null;
                }
                dropdownService.DropdownList.Remove(projectName);
                allStartupProjects.Remove(projectName);
                name2hierarchy.Remove(projectName);
                var projectPath = name2projectPath[projectName];
                name2projectPath.Remove(projectName);
                projectPath2name.Remove(projectPath);
            }
        }

        public void ToggleDebuggingActive(bool debuggingActive)
        {
            logger.LogInfo(debuggingActive ? "Start debugging, disable combobox" : "Stop debugging, enable combobox");
            // When debugging command UI context is activated, disable combobox, otherwise enable combobox
            dropdownService.DropdownEnabled = !debuggingActive;
        }

        private void _LoadConfigurationAndUpdateDropdown()
        {
            configuration = configurationLoader.Load();
            // Check if all manually configured startup projects exist
            if (configuration.MultiProjectConfigurations.Any(config => config.Projects.Any(configProject => !allStartupProjects.Contains(configProject.Name))))
            {
                MessageBox.Show("The configuration file refers to inexistent projects.\nPlease check your configuration file!", "SwitchStartupProject", MessageBoxButton.OK, MessageBoxImage.Exclamation);
            }
            _PopulateDropdownList();
            // Select the currently active startup configuration in the dropdown
            UpdateStartupProject();
        }

        private void _PopulateDropdownList()
        {
            var startupProjects = new List<string>();
            if (configuration.ListAllProjects)
            {
                allStartupProjects.Sort();
                startupProjects.AddRange(allStartupProjects);
            }
            configuration.MultiProjectConfigurations.ForEach(c => startupProjects.Add(c.Name));
            dropdownService.DropdownList = startupProjects;
        }

        private void _RenameProject(IVsHierarchy pHierarchy, string oldName, string oldPath, string newName, string newPath)
        {

            _RenameEntryInList(dropdownService.DropdownList, oldName, newName);
            _RenameEntryInList(allStartupProjects, oldName, newName);

            name2hierarchy.Remove(oldName);
            name2hierarchy.Add(newName, pHierarchy);
            name2projectPath.Remove(oldName);
            name2projectPath.Add(newName, newPath);
            projectPath2name.Remove(oldPath);
            projectPath2name.Add(newPath, newName);

        }

        private void _ChangeStartupProject(string newStartupProject)
        {
            dropdownService.CurrentDropdownValue = newStartupProject;
            if (newStartupProject == null)
            {
                // No startup project
                _ActivateSingleProjectConfiguration(null);
            }
            else if (name2projectPath.ContainsKey(newStartupProject))
            {
                // Single startup project
                _ActivateSingleProjectConfiguration(newStartupProject);
            }
            else if (configuration.MultiProjectConfigurations.Any(c => c.Name == newStartupProject))
            {
                // Multiple startup projects
                var config = configuration.MultiProjectConfigurations.Single(c => c.Name == newStartupProject);
                _ActivateMultiProjectConfiguration(config);
            }
        }

        private MultiProjectConfiguration _GetCurrentlyActiveConfiguration()
        {
            var newStartupProjects = dte.Solution.SolutionBuild.StartupProjects as Array;
            if (newStartupProjects == null) return null;
            if (projectPath2name.Count == 0) return null;

            return new MultiProjectConfiguration(null, newStartupProjects.Cast<string>().Select(projectPath =>
            {
                var projectName = projectPath2name[projectPath];
                var cla = _GetStartArgumentsOfProject(projectName);
                return new MultiProjectConfigurationProject(projectName, cla);
            }).ToList());
        }

        private void _ActivateSingleProjectConfiguration(string projectName)
        {
            _SuspendChangedEvent(() =>
            {
                var projectPath = name2projectPath[projectName];
                dte.Solution.SolutionBuild.StartupProjects = projectPath;
            });
        }

        private void _ActivateMultiProjectConfiguration(MultiProjectConfiguration configuration)
        {
            _SuspendChangedEvent(() =>
            {
                if (configuration.Projects.Count == 1)
                {
                    // If the multi-project startup configuration contains a single project only, handle it as if it was a single-project configuration
                    var projectPath = name2projectPath[configuration.Projects.Single().Name];
                    dte.Solution.SolutionBuild.StartupProjects = projectPath;
                }
                else
                {
                    // SolutionBuild.StartupProjects expects an array of objects
                    var projectPathArray = configuration.Projects.Select(projectConfig => (object)name2projectPath[projectConfig.Name]).ToArray();
                    dte.Solution.SolutionBuild.StartupProjects = projectPathArray;
                }

                // Set CLA
                foreach (var projectConfig in configuration.Projects)
                {
                    _SetStartArgumentsOfProject(projectConfig.Name, projectConfig.CommandLineArguments);
                }
            });
        }

        private string _GetStartArgumentsOfProject(string projectName)
        {
            if (!name2hierarchy.ContainsKey(projectName)) return null;
            var hierarchy = name2hierarchy[projectName];
            var project = projectHierarchyHelper.GetProjectFromHierarchy(hierarchy);
            var configuration = _GetActiveConfigurationOfProject(project);
            var property = _GetStartArgumentsPropertyOfConfiguration(configuration, hierarchy);
            if (property == null) return null;
            return (string)property.Value;
        }

        private void _SetStartArgumentsOfProject(string projectName, string commandLineArguments)
        {
            if (commandLineArguments == null) return;
            if (!name2hierarchy.ContainsKey(projectName)) return;
            var hierarchy = name2hierarchy[projectName];
            var project = projectHierarchyHelper.GetProjectFromHierarchy(hierarchy);
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
            if (properties == null) return null;
            var projectTypeGuids = _GetProjectTypeGuids(projectHierarchy);
            var startArgumentsPropertyName = projectTypeGuids.Contains(GuidList.guidCPlusPlus) ? "CommandArguments" : "StartArguments";
            return properties.Cast<Property>().FirstOrDefault(property => property.Name == startArgumentsPropertyName);
        }

        private void _SuspendChangedEvent(Action action)
        {
            reactToChangedEvent = false;
            action();
            reactToChangedEvent = true;
        }

        private string _GetProjectName(IVsHierarchy pHierarchy)
        {
            object nameObj = null;
            return pHierarchy.GetProperty((uint)VSConstants.VSITEMID.Root, (int)__VSHPROPID.VSHPROPID_Name, out nameObj) == VSConstants.S_OK ?
                (string)nameObj :
                null;
        }

        private string _GetProjectStringProperty(IVsHierarchy pHierarchy, __VSHPROPID property)
        {
            object valueObject = null;
            return pHierarchy.GetProperty((uint)VSConstants.VSITEMID.Root, (int)property, out valueObject) == VSConstants.S_OK ? (string)valueObject : null;
        }

        private Guid? _GetProjectGuidProperty(IVsHierarchy pHierarchy, __VSHPROPID property)
        {
            Guid guid = Guid.Empty;
            return pHierarchy.GetGuidProperty((uint)VSConstants.VSITEMID.Root, (int)property, out guid) == VSConstants.S_OK ? guid : (Guid?)null;
        }

        private IEnumerable<Guid> _GetProjectTypeGuids(IVsHierarchy pHierarchy)
        {
            IEnumerable<Guid> projectTypeGuids;
            var aggregatableProject = pHierarchy as IVsAggregatableProject;
            if (aggregatableProject != null)
            {
                string projectTypeGuidString;
                aggregatableProject.GetAggregateProjectTypeGuids(out projectTypeGuidString);
                projectTypeGuids = projectTypeGuidString.Split(';')
                    .Where(guidString => !string.IsNullOrEmpty(guidString))
                    .Select(guidString => new Guid(guidString));
            }
            else
            {
                var project = projectHierarchyHelper.GetProjectFromHierarchy(pHierarchy);
                projectTypeGuids = new[] { new Guid(project.Kind) };
            }
            return projectTypeGuids;
        }

        private string _GetPathRelativeToSolution(IVsHierarchy pHierarchy)
        {
            var fullProjectPath = _GetAbsolutePath(pHierarchy);
            var solutionPath = Path.GetDirectoryName(dte.Solution.FullName) + @"\";
            return Paths.GetPathRelativeTo(fullProjectPath, solutionPath);
        }

        private string _GetAbsolutePath(IVsHierarchy pHierarchy)
        {
            var project = projectHierarchyHelper.GetProjectFromHierarchy(pHierarchy);
            return project.FullName;
        }


        private void _AddProject(string name, IVsHierarchy pHierarchy, bool isWebSiteProject)
        {
            name2hierarchy.Add(name, pHierarchy);
            // Website projects need to be set using the full path
            name2projectPath.Add(name, isWebSiteProject ? _GetAbsolutePath(pHierarchy) : _GetPathRelativeToSolution(pHierarchy));
            projectPath2name.Add(isWebSiteProject ? _GetAbsolutePath(pHierarchy) : _GetPathRelativeToSolution(pHierarchy), name);
        }

        private void _ClearProjects()
        {
            name2hierarchy.Clear();
            name2projectPath.Clear();
            projectPath2name.Clear();
            allStartupProjects = new List<string>();
            dropdownService.DropdownList = null;
        }

        private void _RenameEntryInList(IList<string> list, string oldName, string newName)
        {
            for (int i = 0; i < list.Count; i++)
            {
                if (list[i] == oldName)
                {
                    list[i] = newName;
                }
            }
        }
    }
}
