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
        private readonly OptionPage options;
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



        public StartupProjectSwitcher(OleMenuCommand combobox, OptionPage options, DTE dte, IVsFileChangeEx fileChangeService, ProjectHierarchyHelper project2Hierarchy, Package package, int mruCount, SwitchStartupProjectPackage.ActivityLogger logger)
        {
            logger.LogInfo("Entering constructor of StartupProjectSwitcher");
            this.menuSwitchStartupProjectComboCommand = combobox;
            this.options = options;
            this.dte = dte;
            this.fileChangeService = fileChangeService;
            this.projectHierarchyHelper = project2Hierarchy;
            this.openOptionsPage = () => package.ShowOptionPage(typeof(OptionPage));
            this.logger = logger;

            // initialize MRU list
            mruStartupProjects = new MRUList<string>(mruCount);

            // initialize type list
            typeStartupProjects = new List<string>();

            // initialize all list
            allStartupProjects = new List<string>();

            multiProjectConfigurations = new List<MultiProjectConfiguration>();

            options.Modified += (s, e) =>
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
            options.GetAllProjectNames = () => allStartupProjects;
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
            if (openingSolution || !options.EnableMultiProjectConfiguration) return;

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
            _LoadSettings();
            // When solution is open: enable multi-project configuration
            logger.LogInfo("Enable multi project configuration");
            options.EnableMultiProjectConfiguration = true;
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
            options.Configurations.Clear();
            logger.LogInfo("Disable multi project configuration");
            options.EnableMultiProjectConfiguration = false;
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
            bool valid = true;
            object nameObj = null;
            object typeNameObj = null;
            object captionObj = null;
            Guid guid = Guid.Empty;
            valid &= pHierarchy.GetProperty((uint)VSConstants.VSITEMID.Root, (int)__VSHPROPID.VSHPROPID_Name, out nameObj) == VSConstants.S_OK;
            valid &= pHierarchy.GetProperty((uint)VSConstants.VSITEMID.Root, (int)__VSHPROPID.VSHPROPID_TypeName, out typeNameObj) == VSConstants.S_OK;
            valid &= pHierarchy.GetProperty((uint)VSConstants.VSITEMID.Root, (int)__VSHPROPID.VSHPROPID_Caption, out captionObj) == VSConstants.S_OK;
            valid &= pHierarchy.GetGuidProperty((uint)VSConstants.VSITEMID.Root, (int)__VSHPROPID.VSHPROPID_TypeGuid, out guid) == VSConstants.S_OK;
            
            if (valid)
            {
                var name = (string)nameObj;
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

        private void _LoadSettings()
        {
            logger.LogInfo("Loading configuration for solution");
            mruStartupProjects = new MRUList<string>(options.MostRecentlyUsedCount, settingsPersister.GetSingleProjectMruList().Intersect(allStartupProjects));
            options.Configurations.Clear();
            settingsPersister.GetMultiProjectConfigurations().ForEach(options.Configurations.Add);
            options.ActivateCommandLineArguments = settingsPersister.GetActivateCommandLineArguments();
        }

        private void _StoreSettings()
        {
            logger.LogInfo("Storing solution specific configuration.");
            settingsPersister.StoreSingleProjectMruList(mruStartupProjects);
            settingsPersister.StoreMultiProjectConfigurations(options.Configurations);
            settingsPersister.StoreActivateCommandLineArguments(options.ActivateCommandLineArguments);
            settingsPersister.Persist();
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
                if (options.ActivateCommandLineArguments)
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
                if (options.ActivateCommandLineArguments)
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

        private void _PopulateStartupProjects()
        {
            switch (options.Mode)
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
                    currentStartupProject = options.Configurations[index].Name;
                }
            }

            multiProjectConfigurations.Clear();
            multiProjectConfigurations = (from configuration in options.Configurations
                                          let projects = (from project in configuration.Projects
                                                          where project.Name != null && name2projectPath.ContainsKey(project.Name)
                                                          select new MultiProjectConfigurationProject(project.Name, project.CommandLineArguments)).ToList()
                                          select new MultiProjectConfiguration(configuration.Name, projects)).ToList();

        }

        private void _ChangeMRUCountInOptions()
        {
            var oldList = mruStartupProjects;
            mruStartupProjects = new MRUList<string>(options.MostRecentlyUsedCount, oldList);
            _PopulateStartupProjects();
        }
    }
}
