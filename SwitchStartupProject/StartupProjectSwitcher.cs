using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Text;
using EnvDTE;

using Microsoft.VisualStudio;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;

namespace LucidConcepts.SwitchStartupProject
{
    public class StartupProjectSwitcher
    {
        private readonly SolutionOptionPage solutionOptions;
        private readonly OptionPage defaultOptions;
        private readonly OleMenuCommand menuSwitchStartupProjectComboCommand;
        private readonly Action openOptionsPage;
        private readonly SwitchStartupProjectPackage.ActivityLogger logger;

        private readonly Dictionary<string, IVsHierarchy> name2hierarchy = new Dictionary<string, IVsHierarchy>();
        private readonly Dictionary<string, string> name2projectPath = new Dictionary<string, string>();
        private readonly Dictionary<string, string> projectPath2name = new Dictionary<string, string>();

        private ConfigurationsPersister settingsPersister;
        private MRUList<string> mruStartupProjects;
        private List<string> typeStartupProjects;
        private List<string> allStartupProjects;
        private List<MultiProjectConfiguration> multiProjectConfigurations;

        private const string sentinel = "";
        private const string configure = "Configure...";
        private List<string> startupProjects = new List<string>(new [] { configure });
        private string currentStartupProject = sentinel;
        private bool openingSolution;
        private bool reactToChangedEvent = true;

        private readonly DTE dte;
        private readonly IVsFileChangeEx fileChangeService;
        private readonly ProjectHierarchyHelper projectHierarchyHelper;



        public StartupProjectSwitcher(OleMenuCommand combobox, SolutionOptionPage solutionOptions, OptionPage defaultOptions, DTE dte, IVsFileChangeEx fileChangeService, ProjectHierarchyHelper project2Hierarchy, Package package, int mruCount, SwitchStartupProjectPackage.ActivityLogger logger)
        {
            logger.LogInfo("Entering constructor of StartupProjectSwitcher");
            this.menuSwitchStartupProjectComboCommand = combobox;
            this.solutionOptions = solutionOptions;
            this.defaultOptions = defaultOptions;
            this.dte = dte;
            this.fileChangeService = fileChangeService;
            this.projectHierarchyHelper = project2Hierarchy;
            this.openOptionsPage = () => package.ShowOptionPage(solutionOptions.EnableSolutionConfiguration ? typeof(SolutionOptionPage) : typeof(OptionPage));
            this.logger = logger;

            // initialize MRU list
            mruStartupProjects = new MRUList<string>(mruCount);

            // initialize type list
            typeStartupProjects = new List<string>();

            // initialize all list
            allStartupProjects = new List<string>();

            multiProjectConfigurations = new List<MultiProjectConfiguration>();

            solutionOptions.Modified += (s, e) =>
            {
                if (e.OptionParameter == EOptionParameter.Mode)
                {
                    _PopulateStartupProjects();
                }
                else if (e.OptionParameter == EOptionParameter.MruCount)
                {
                    _ChangeMRUCountInOptions();
                }
                else if (e.OptionParameter == EOptionParameter.MultiProjectConfigurations)
                {
                    _ChangeMultiProjectConfigurationsInOptions(e.ListChangedEventArgs);
                    _PopulateStartupProjects();
                }
            };
            solutionOptions.GetAllProjectNames = () => allStartupProjects;
            menuSwitchStartupProjectComboCommand.Enabled = true;
        }

        public string GetCurrentStartupProject()
        {
            return currentStartupProject;
        }

        public Array GetStartupProjectChoices()
        {
            return startupProjects.ToArray();
        }

        public void ChooseStartupProject(string name)
        {
            // new value was selected or typed in
            // see if it is the configuration item
            if (String.Compare(configure, name, StringComparison.CurrentCultureIgnoreCase) == 0)
            {
                logger.LogInfo("Selected configure... in combobox");
                openOptionsPage();
                return;

            }

            // see if it is one of our items
            foreach (string project in this.startupProjects)
            {
                if (String.Compare(project, name, StringComparison.CurrentCultureIgnoreCase) == 0)
                {
                    logger.LogInfo("Selected or typed new entry in combobox: {0}", project);
                    _ChangeStartupProject(project);
                    return;
                }
            }
            throw (new ArgumentException("ParamNotValidStringInList")); // force an exception to be thrown
        }

        public void UpdateStartupProject()
        {
            // When startup project is set through dropdown, don't do anything
            if (!reactToChangedEvent) return;
            // Don't react to startup project changes while opening the solution or when multi-project configurations have not yet been loaded.
            if (openingSolution || !solutionOptions.EnableSolutionConfiguration) return;

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
                    mruStartupProjects.Touch(newStartupProjectName);
                    _PopulateStartupProjects();
                    currentStartupProject = newStartupProjectName;
                    return;
                }
                logger.LogInfo("New unknown startup project was activated outside of combobox");
                currentStartupProject = sentinel;
                return;
            }

            var newConfig = _GetCurrentlyActiveConfiguration();
            _SelectMultiProjectConfigInDropdown(newConfig,
                success: newStartupProjectName => logger.LogInfo("New multi-project startup config was activated outside of combobox: {0}", newStartupProjectName),
                failure: () =>
                {
                    logger.LogInfo("New unknown multi-project startup config was activated outside of combobox");
                    currentStartupProject = sentinel;
                });
        }

        private void _SelectMultiProjectConfigInDropdown(MultiProjectConfiguration newStartupProjects, Action<string> success, Action failure)
        {
            foreach (var configuration in multiProjectConfigurations.Where(configuration => _AreEqual(configuration.Projects, newStartupProjects.Projects)))
            {
                var newStartupProjectName = configuration.Name;
                currentStartupProject = newStartupProjectName;
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
            settingsPersister = new ConfigurationsPersister(dte.Solution.FullName, ".startup.suo", fileChangeService);
            settingsPersister.SettingsFileModified += OnSettingsFileModified;
            if (settingsPersister.ConfigurationFileExists())
            {
                _LoadSettings();
            }
            else
            {
                _LoadDefaultSettings();
            }
            // When solution is open: enable solution configuration
            logger.LogInfo("Enable solution configuration");
            solutionOptions.EnableSolutionConfiguration = true;
            _PopulateStartupProjects();
            // Select the currently active startup configuration in the dropdown
            UpdateStartupProject();
        }

        public void BeforeCloseSolution()
        {
            logger.LogInfo("Starting to close solution");
            if (settingsPersister == null)  // This happens e.g. when creating a new website
            {
                logger.LogInfo("Cannot store solution specific configuration because configuration persister has not yet been set up.");
                return;
            }
            _StoreSettings();
            settingsPersister.SettingsFileModified -= OnSettingsFileModified;
        }

        public void AfterCloseSolution()
        {
            logger.LogInfo("Finished to close solution");
            // When solution is closed: choose no project
            currentStartupProject = sentinel;
            solutionOptions.Configurations.Clear();
            logger.LogInfo("Disable solution configuration");
            solutionOptions.EnableSolutionConfiguration = false;
            settingsPersister = null;
            _ClearProjects();
        }

        public void OnSettingsFileModified(object sender, SettingsFileModifiedEventArgs args)
        {
            settingsPersister.Load();
            _LoadSettings();
        }

        public void OnSolutionSaved()
        {
            if (settingsPersister == null)  // This happens e.g. when creating a new website
            {
                logger.LogInfo("Cannot store solution specific configuration because configuration persister has not yet been set up.");
                return;
            }
            _StoreSettings();
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

            var project = projectHierarchyHelper.GetProjectFromHierarchy(pHierarchy);
            var projectTypeGuids = _GetProjectTypeGuids(pHierarchy);

            var isWebSiteProject = projectTypeGuids.Contains(GuidList.guidWebSite);

            _AddProject(name, pHierarchy, isWebSiteProject);
            allStartupProjects.Add(name);

            VSLangProj.prjOutputType? projectOutputType = null;
            if (!projectTypeGuids.Contains(GuidList.guidCPlusPlus))
            {
                try
                {
                    var outputType = project.Properties.Item("OutputType");
                    if (outputType != null)
                    {
                        if (outputType.Value is int)
                        {
                            projectOutputType = (VSLangProj.prjOutputType) outputType.Value;
                        }
                        else if (outputType.Value.GetType().FullName == "Microsoft.VisualStudio.FSharp.ProjectSystem.OutputType")
                        {
                            switch (outputType.Value.ToString())
                            {
                                case "WinExe":
                                    projectOutputType = VSLangProj.prjOutputType.prjOutputTypeWinExe;
                                    break;
                        
                                case "Exe":
                                    projectOutputType = VSLangProj.prjOutputType.prjOutputTypeExe;
                                    break;
                            }
                        }
                    }
                }
                catch (Exception)
                {
                }
            }

            // Smart mode: Add executables, windows executables, web applications, Visual Studio extension packages and database projects
            if (projectOutputType == VSLangProj.prjOutputType.prjOutputTypeExe ||
                projectOutputType == VSLangProj.prjOutputType.prjOutputTypeWinExe ||
                projectOutputType == VSLangProj.prjOutputType.prjOutputTypeLibrary && (
                    projectTypeGuids.Contains(GuidList.guidWebApp) ||
                    projectTypeGuids.Contains(GuidList.guidVsPackage)) ||
                projectOutputType == null && projectTypeGuids.Contains(GuidList.guidDatabase))
            {
                typeStartupProjects.Add(name);
            }

            if (!openingSolution)
            {
                _PopulateStartupProjects(); // when reopening a single project, refresh list
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
            var reselectRenamedProject = currentStartupProject == oldName;
            _RenameProject(pHierarchy, oldName, oldPath, newName, newPath);
            if (reselectRenamedProject)
            {
                currentStartupProject = newName;
            }
        }

        public void CloseProject(IVsHierarchy pHierarchy, string projectName)
        {
            logger.LogInfo("Closing project: {0}", projectName);
            // When project is closed: remove it from list of startup projects (if it was in there)
            if (allStartupProjects.Contains(projectName))
            {
                if (currentStartupProject == projectName)
                {
                    currentStartupProject = sentinel;
                }
                startupProjects.Remove(projectName);
                allStartupProjects.Remove(projectName);
                typeStartupProjects.Remove(projectName);
                mruStartupProjects.Remove(projectName);
                name2hierarchy.Remove(projectName);
                var projectPath = name2projectPath[projectName];
                name2projectPath.Remove(projectName);
                projectPath2name.Remove(projectPath);
            }
        }

        public void ToggleDebuggingActive(bool active)
        {
            logger.LogInfo(active ? "Start debugging, disable combobox" : "Stop debugging, enable combobox");
            // When debugging command UI context is activated, disable combobox, otherwise enable combobox
            menuSwitchStartupProjectComboCommand.Enabled = !active;
        }

        private void _LoadDefaultSettings()
        {
            logger.LogInfo("Loading default configuration");
            solutionOptions.Mode = defaultOptions.Mode;
            solutionOptions.MostRecentlyUsedCount = defaultOptions.MostRecentlyUsedCount;
            mruStartupProjects = new MRUList<string>(solutionOptions.MostRecentlyUsedCount);
            solutionOptions.Configurations.Clear();
            solutionOptions.ActivateCommandLineArguments = false;
        }

        private void _LoadSettings()
        {
            logger.LogInfo("Loading configuration for solution");
            solutionOptions.Mode = settingsPersister.GetSingleProjectMode();
            solutionOptions.MostRecentlyUsedCount = settingsPersister.GetSingleProjectMruCount();
            mruStartupProjects = new MRUList<string>(solutionOptions.MostRecentlyUsedCount, settingsPersister.GetSingleProjectMruList().Intersect(allStartupProjects));
            solutionOptions.Configurations.Clear();
            settingsPersister.GetMultiProjectConfigurations().ForEach(solutionOptions.Configurations.Add);
            solutionOptions.ActivateCommandLineArguments = settingsPersister.GetActivateCommandLineArguments();
        }

        private void _StoreSettings()
        {
            logger.LogInfo("Storing solution specific configuration.");
            settingsPersister.StoreSingleProjectMode(solutionOptions.Mode);
            settingsPersister.StoreSingleProjectMruCount(solutionOptions.MostRecentlyUsedCount);
            settingsPersister.StoreSingleProjectMruList(mruStartupProjects);
            settingsPersister.StoreMultiProjectConfigurations(solutionOptions.Configurations);
            settingsPersister.StoreActivateCommandLineArguments(solutionOptions.ActivateCommandLineArguments);
            settingsPersister.Persist();
        }

        private void _RenameProject(IVsHierarchy pHierarchy, string oldName, string oldPath, string newName, string newPath)
        {

            _RenameEntryInList(startupProjects, oldName, newName);
            _RenameEntryInList(allStartupProjects, oldName, newName);
            _RenameEntryInList(typeStartupProjects, oldName, newName);
            mruStartupProjects.Replace(oldName, newName);

            name2hierarchy.Remove(oldName);
            name2hierarchy.Add(newName, pHierarchy);
            name2projectPath.Remove(oldName);
            name2projectPath.Add(newName, newPath);
            projectPath2name.Remove(oldPath);
            projectPath2name.Add(newPath, newName);


            solutionOptions.Configurations.ForEach(config => config.Projects.ForEach(project =>
            {
                if (project.Name == oldName)
                {
                    project.Name = newName;
                }
            }));
        }

        private void _ChangeStartupProject(string newStartupProject)
        {
            this.currentStartupProject = newStartupProject;
            if (newStartupProject == sentinel)
            {
                // No startup project
                _ActivateSingleProjectConfiguration(null);
            }
            else if (name2projectPath.ContainsKey(newStartupProject))
            {
                // Single startup project
                _ActivateSingleProjectConfiguration(newStartupProject);
            }
            else if (multiProjectConfigurations.Any(c => c.Name == newStartupProject))
            {
                // Multiple startup projects
                var configuration = multiProjectConfigurations.Single(c => c.Name == newStartupProject);
                _ActivateMultiProjectConfiguration(configuration);
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
                
                // Clear CLA
                if (solutionOptions.ActivateCommandLineArguments)
                {
                    _SetStartArgumentsOfProject(projectName, string.Empty);
                }
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
                if (solutionOptions.ActivateCommandLineArguments)
                {
                    foreach (var projectConfig in configuration.Projects)
                    {
                        _SetStartArgumentsOfProject(projectConfig.Name, projectConfig.CommandLineArguments);
                    }
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

        private Configuration _GetActiveConfigurationOfProject(Project project)
        {
            if (project == null) return null;
            var configurationManager = project.ConfigurationManager;
            return configurationManager == null ? null : configurationManager.ActiveConfiguration;
        }

        private IEnumerable<Configuration> _GetAllConfigurationsOfProject(Project project)
        {
            if (project == null) return null;
            var configurationManager = project.ConfigurationManager;
            return configurationManager.Cast<Configuration>();
        }

        private Property _GetStartArgumentsPropertyOfConfiguration(Configuration configuration, IVsHierarchy projectHierarchy)
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
            typeStartupProjects = new List<string>();
            startupProjects = new List<string>(new [] { configure });
        }

        private void _RenameEntryInList(List<string> list, string oldName, string newName)
        {
            var index = list.FindIndex(p => p == oldName);
            if (index < 0) return;
            list[index] = newName;
        }

        private void _PopulateStartupProjects()
        {
            switch (solutionOptions.Mode)
            {
                case EMode.All:
                    allStartupProjects.Sort();
                    startupProjects = allStartupProjects.ToList();
                    break;
                case EMode.MostRecentlyUsed:
                    startupProjects = mruStartupProjects.ToList();
                    break;
                case EMode.Smart:
                    typeStartupProjects.Sort();
                    startupProjects = typeStartupProjects.ToList();
                    break;
                case EMode.None:
                    startupProjects = new List<string>();
                    break;
            }
            multiProjectConfigurations.ForEach(c => startupProjects.Add(c.Name));
            startupProjects.Add(configure);
        }

        private void _ChangeMultiProjectConfigurationsInOptions(ListChangedEventArgs listChangedEventArgs)
        {
            // If the currently selected multi-project startup configuration is changed, adjust its name
            if (listChangedEventArgs != null &&
                listChangedEventArgs.ListChangedType == ListChangedType.ItemChanged)
            {
                var index = listChangedEventArgs.NewIndex;
                if (index < multiProjectConfigurations.Count &&
                    multiProjectConfigurations[index].Name == currentStartupProject)
                {
                    currentStartupProject = solutionOptions.Configurations[index].Name;
                }
            }

            multiProjectConfigurations.Clear();
            multiProjectConfigurations = (from configuration in solutionOptions.Configurations
                                          let projects = (from project in configuration.Projects
                                                          where project.Name != null && name2projectPath.ContainsKey(project.Name)
                                                          select new MultiProjectConfigurationProject(project.Name, project.CommandLineArguments)).ToList()
                                          select new MultiProjectConfiguration(configuration.Name, projects)).ToList();

        }

        private void _ChangeMRUCountInOptions()
        {
            var oldList = mruStartupProjects;
            mruStartupProjects = new MRUList<string>(solutionOptions.MostRecentlyUsedCount, oldList);
            _PopulateStartupProjects();
        }
    }
}
