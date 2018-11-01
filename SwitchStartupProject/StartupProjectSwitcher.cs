using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows;

using EnvDTE;

using Microsoft.VisualStudio;
using Microsoft.VisualStudio.ProjectSystem.Properties;
using Microsoft.VisualStudio.ProjectSystem.Debug;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;

using Newtonsoft.Json;

using Task = System.Threading.Tasks.Task;

namespace LucidConcepts.SwitchStartupProject
{
    public class StartupProjectSwitcher
    {
        private readonly DropdownService dropdownService;

        private Solution solution = null;
        private bool reactToChangedEvent = true;

        private readonly DTE dte;
        private readonly IVsFileChangeEx fileChangeService;

        public StartupProjectSwitcher(DropdownService dropdownService, DTE dte, IVsFileChangeEx fileChangeService)
        {
            this.dropdownService = dropdownService;
            dropdownService.OnListItemSelectedAsync = entry => _ChangeStartupProjectAsync(entry, store: true);
            dropdownService.OnConfigurationSelected = _ShowMsgOpenSolution;
            this.dte = dte;
            this.fileChangeService = fileChangeService;
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

            // If new startup project(s) are the same as the already selected do nothing
            if (dropdownService.CurrentDropdownValue is SingleProjectDropdownEntry singleEntry && singleEntry.MatchesPaths(newStartupProjectPaths)) return;
            if (dropdownService.CurrentDropdownValue is MultiProjectDropdownEntry multiEntry && multiEntry.MatchesPaths(newStartupProjectPaths)) return;

            // Try to find a matching dropdown entry
            var bestMatch = dropdownService.DropdownList.OfType<SingleProjectDropdownEntry>().SingleOrDefault(entry => entry.MatchesPaths(newStartupProjectPaths)) ??
                (IDropdownEntry) dropdownService.DropdownList.OfType<MultiProjectDropdownEntry>().FirstOrDefault(entry => entry.MatchesPaths(newStartupProjectPaths));
            if (bestMatch != null)
            {
                Logger.Log("New startup configuration was activated outside of dropdown: {0}", bestMatch.DisplayName);
                dropdownService.CurrentDropdownValue = bestMatch;
                return;
            }

            Logger.Log("Unknown startup configuration was activated outside of dropdown");
            dropdownService.CurrentDropdownValue = DropdownService.OtherItem;
        }

        // Is called before a solution and its projects are loaded.
        // Is NOT called when a new solution and project are created.
        public void BeforeOpenSolution(string solutionFileName)
        {
            solution = new Solution
            {
                IsOpening = true
            };
        }

        // Is called after a solution and its projects have been loaded.
        // Is also called when a new solution and project have been created.
        public void AfterOpenSolution()
        {
            if (solution == null)   // This happens when creating a new solution
            {
                Logger.Log("Created a new solution");
                solution = new Solution();
            }
            solution.IsOpening = false;
            if (string.IsNullOrEmpty(dte.Solution.FullName))  // This happens e.g. when creating a new website
            {
                Logger.Log("Solution path not yet known. Skipping initialization of configuration persister and loading of settings.");
                return;
            }
            var configurationFilename = ConfigurationLoader.GetConfigurationFilename(dte.Solution.FullName);
            var oldConfigurationFilename = ConfigurationLoader.GetOldConfigurationFilename(dte.Solution.FullName);
            solution.ConfigurationLoader = new ConfigurationLoader(configurationFilename);
            solution.ConfigurationFileTracker = new ConfigurationFileTracker(configurationFilename, fileChangeService, _LoadConfigurationAndUpdateSettingsOfCurrentStartupProject);
            var configurationFileOpener = new ConfigurationFileOpener(dte, configurationFilename, oldConfigurationFilename, solution.ConfigurationLoader);
            dropdownService.OnConfigurationSelected = configurationFileOpener.Open;
            var activeConfigurationFilename = ActiveConfigurationLoader.GetConfigurationFilename(dte.Solution.FullName);
            solution.ActiveConfigurationLoader = new ActiveConfigurationLoader(activeConfigurationFilename);

            _LoadConfigurationAndPopulateDropdown();
            // Determine the last active startup configuration and select it in the dropdown
            dropdownService.CurrentDropdownValue = solution.ActiveConfigurationLoader.Load(dropdownService.DropdownList);
        }

        public void BeforeCloseSolution()
        {
            if (solution == null || solution.ConfigurationFileTracker == null)
            {
                return;
            }
            solution.ConfigurationFileTracker.Stop();
            solution.ConfigurationFileTracker = null;
        }

        public void AfterCloseSolution()
        {
            // When solution is closed: choose no project
            dropdownService.OnConfigurationSelected = _ShowMsgOpenSolution;
            dropdownService.CurrentDropdownValue = null;
            solution = null;
            dropdownService.DropdownList = null;
        }

        public void OpenProject(IVsHierarchy pHierarchy, bool isCreated)
        {
            // When project is opened: register it and its name

            // Don't register project if solution path is not yet available
            if (string.IsNullOrEmpty(dte.Solution.FullName)) return;

            // Filter out hierarchy elements that don't represent projects
            var project = SolutionProject.FromHierarchy(pHierarchy, dte.Solution.FullName);
            if (project == null) return;

            if (solution == null)   // This happens e.g. when creating a new project or website solution
            {
                solution = new Solution
                {
                    IsOpening = true,
                };
            }
            if (!solution.IsOpening)
            {
                _PopulateDropdownList(); // when reopening a single project, refresh list
            }
        }

        public void RenameProject(IVsHierarchy pHierarchy)
        {
            _PopulateDropdownList();
        }

        public void CloseProject(IVsHierarchy pHierarchy)
        {
            // When project is closed: remove it from list of startup projects
            //if ((dropdownService.CurrentDropdownValue is SingleProjectDropdownEntry) &&
            //    (dropdownService.CurrentDropdownValue as SingleProjectDropdownEntry).Project == project)
            //{
            //    dropdownService.CurrentDropdownValue = null;
            //}
            _PopulateDropdownList();
        }

        public void ToggleDebuggingActive(bool debuggingActive)
        {
            // When debugging command UI context is activated, disable combobox, otherwise enable combobox
            dropdownService.DropdownEnabled = !debuggingActive;
        }

        private IList<SolutionProject> _GetProjects()
        {
            return (from hierarchy in _GetProjectHierarchies()
                let project = SolutionProject.FromHierarchy(hierarchy, dte.Solution.FullName)
                where project != null
                select project).ToArray();
        }

        private IEnumerable<IVsHierarchy> _GetProjectHierarchies()
        {
            var solution = ServiceProvider.GlobalProvider.GetService(typeof(SVsSolution)) as IVsSolution2;
            if (solution == null) yield break;

            IEnumHierarchies enumerator = null;
            Guid guid = Guid.Empty;
            solution.GetProjectEnum((uint)__VSENUMPROJFLAGS.EPF_LOADEDINSOLUTION, ref guid, out enumerator);
            IVsHierarchy[] hierarchy = new IVsHierarchy[1] { null };
            uint fetched = 0;
            enumerator.Reset();
            while (enumerator.Next(1, hierarchy, out fetched) == VSConstants.S_OK && fetched == 1)
            {
                yield return hierarchy[0];
            }
        }

        private async Task _LoadConfigurationAndUpdateSettingsOfCurrentStartupProject()
        {
            var selectedDropdownValue = dropdownService.CurrentDropdownValue;
            _LoadConfigurationAndPopulateDropdown();
            var equivalentItem = dropdownService.DropdownList.FirstOrDefault(item => item.IsEqual(selectedDropdownValue));
            if (equivalentItem == null) return;
            await _ChangeStartupProjectAsync(equivalentItem, store: false);
        }

        private void _LoadConfigurationAndPopulateDropdown()
        {
            solution.Configuration = solution.ConfigurationLoader.Load();
            // Check if any manually configured startup project references need to be disambiguated
            var projectsByName = _GetProjects().ToLookup(project => project.Name);
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
                    Logger.LogActive("\nWARNING: The configuration file refers to ambiguous project names.");
                    Logger.Log("Please use either of the following project paths instead:\n\n" + projectNameToPath);
                }
            }
            _PopulateDropdownList();
        }

        private void _PopulateDropdownList()
        {
            var startupProjects = new List<IDropdownEntry>();
            if (solution != null)   // Solution may be null e.g. when creating a new website
            {
                var projects = _GetProjects();
                if (solution.Configuration.ListAllProjects)
                {
                    var projectsByName = projects.ToLookup(project => project.Name);
                    projects.OrderBy(project => project.Name).ForEach(project => startupProjects.Add(new SingleProjectDropdownEntry(project)
                    {
                        Disambiguate = projectsByName[project.Name].Count() > 1
                    }));

                }
                solution.Configuration.MultiProjectConfigurations.ForEach(config => startupProjects.Add(new MultiProjectDropdownEntry(_GetStartupConfiguration(config, projects))));
            }
            dropdownService.DropdownList = startupProjects;
        }

        private StartupConfiguration _GetStartupConfiguration(MultiProjectConfiguration config, IList<SolutionProject> projects)
        {
            var startupConfigurationProjects = from configProject in config.Projects
                                               let project = projects.FirstOrDefault(solutionProject => _ConfigRefersToProject(configProject, solutionProject))
                                               where project != null
                                               select new StartupConfigurationProject(
                                                   project,
                                                   configProject.CommandLineArguments,
                                                   configProject.WorkingDirectory,
                                                   configProject.StartProject,
                                                   configProject.StartExternalProgram,
                                                   configProject.StartBrowserWithUrl,
                                                   configProject.EnableRemoteDebugging,
                                                   configProject.RemoteDebuggingMachine,
                                                   configProject.ProfileName);
            return new StartupConfiguration(config.Name, startupConfigurationProjects.ToList());
        }

        private async Task _ChangeStartupProjectAsync(IDropdownEntry newStartupProject, bool store)
        {
            dropdownService.CurrentDropdownValue = newStartupProject;
            if (newStartupProject == null)
            {
                // No startup project
                _ActivateSingleProjectConfiguration(null);
                return;
            }
            if (store) solution.ActiveConfigurationLoader.Save(newStartupProject);
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
                await ActivateMultiProjectConfigurationAsync(config);
            }
        }

        private bool IsCpsProject(IVsHierarchy hierarchy)
        {
            if (hierarchy == null) return false;
            return hierarchy.IsCapabilityMatch("CPS");
        }

        private bool IsProjectWithLaunchSettings(IVsHierarchy hierarchy)
        {
            if (hierarchy == null) return false;
            return hierarchy.IsCapabilityMatch("LaunchProfiles");
        }

        private void _ActivateSingleProjectConfiguration(SolutionProject project)
        {
            _SuspendChangedEvent(() =>
            {
                dte.Solution.SolutionBuild.StartupProjects = project != null ? project.Path : null;
            });
        }

        private async Task ActivateMultiProjectConfigurationAsync(StartupConfiguration startupConfiguration)
        {
            var startupProjects = startupConfiguration.Projects;
            await _SuspendChangedEventAsync(async () =>
            {
                if (startupProjects.Count == 1)
                {
                    // If the multi-project startup configuration contains a single project only, handle it as if it was a single-project configuration
                    var project = startupProjects.Single().Project;
                    if (project == null)
                    {
                        Logger.LogActive("\nERROR: The activated configuration refers to an inexistent project.");
                        Logger.Log("Please check your configuration file!");
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
                        Logger.LogActive("\nERROR: The activated configuration refers to inexistent projects.");
                        Logger.Log("Please check your configuration file!");
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
                            if (IsProjectWithLaunchSettings(solutionProject.Hierarchy))
                            {
                                var launchSettingsProvider = context.UnconfiguredProject.Services.ExportProvider.GetExportedValue<ILaunchSettingsProvider>();
                                if (launchSettingsProvider != null)
                                {
                                    var launchProfile = startupProject.ProfileName != null ?
                                        launchSettingsProvider.CurrentSnapshot?.Profiles.SingleOrDefault(profile => profile.Name == startupProject.ProfileName) :
                                        launchSettingsProvider.ActiveProfile;
                                    if (launchProfile != null)
                                    {
                                        var writableLaunchProfile = new WritableLaunchProfile(launchProfile);
                                        if (startupProject.CommandLineArguments != null)
                                        {
                                            writableLaunchProfile.CommandLineArgs = startupProject.CommandLineArguments;
                                        }
                                        if (startupProject.WorkingDirectory != null)
                                        {
                                            writableLaunchProfile.WorkingDirectory = startupProject.WorkingDirectory;
                                        }
                                        if (startupProject.StartExternalProgram != null)
                                        {
                                            writableLaunchProfile.ExecutablePath = startupProject.StartExternalProgram;
                                        }
                                        if (startupProject.StartBrowserWithUrl != null)
                                        {
                                            writableLaunchProfile.LaunchUrl = startupProject.StartBrowserWithUrl;
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

                                        await launchSettingsProvider.AddOrUpdateProfileAsync(writableLaunchProfile, false);
                                    }
                                    if (startupProject.ProfileName != null)
                                    {
                                        await launchSettingsProvider.SetActiveProfileAsync(startupProject.ProfileName);
                                    }
                                }
                            }
                        }
                    }
                    // Handle VC++ projects in a special way
                    else if (new Guid(project.Kind) == GuidList.guidCPlusPlus)
                    {
                        var vcProject = (dynamic)project.Object;
                        foreach (var vcConfiguration in vcProject.Configurations)
                        {
                            var vcDebugSettings = vcConfiguration.DebugSettings;
                            _SetPropertyValue(v => vcDebugSettings.CommandArguments = v, startupProject.CommandLineArguments);
                            _SetPropertyValue(v => vcDebugSettings.WorkingDirectory = v, startupProject.WorkingDirectory);
                            _SetPropertyValue(v =>
                            {
                                if (startupProject.EnableRemoteDebugging == true) vcDebugSettings.RemoteCommand = v;
                                else vcDebugSettings.Command = v;
                            }, startupProject.StartExternalProgram);
                            _SetPropertyValue(v => vcDebugSettings.HttpUrl = v, startupProject.StartBrowserWithUrl);
                            _SetPropertyValue(v => vcDebugSettings.RemoteMachine = v, startupProject.RemoteDebuggingMachine);

                            if (startupProject.EnableRemoteDebugging == true) _SetDynamicDebuggerFlavor(vcDebugSettings, "eRemoteDebugger");
                            else if (!string.IsNullOrEmpty(startupProject.StartBrowserWithUrl)) _SetDynamicDebuggerFlavor(vcDebugSettings, "eWebBrowserDebugger");
                            else if (startupProject.EnableRemoteDebugging == false) _SetDynamicDebuggerFlavor(vcDebugSettings, "eLocalDebugger");
                            else if (!string.IsNullOrEmpty(startupProject.StartExternalProgram)) _SetDynamicDebuggerFlavor(vcDebugSettings, "eLocalDebugger");
                            else if (startupProject.StartProject == true) _SetDynamicDebuggerFlavor(vcDebugSettings, "eLocalDebugger");
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

        private void _SetDynamicDebuggerFlavor(dynamic vcDebugSettings, string enumValueName)
        {
            // This is a workaround for enum eDebuggerTypes used with dynamic types.
            vcDebugSettings.DebuggerFlavor = (dynamic)Enum.Parse(vcDebugSettings.DebuggerFlavor.GetType(), enumValueName);
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

        private async Task _SuspendChangedEventAsync(Func<Task> action)
        {
            reactToChangedEvent = false;
            await action();
            reactToChangedEvent = true;
        }

        private void _ShowMsgOpenSolution()
        {
            Logger.LogActive("\nERROR: Please open a solution before you configure its startup projects.");
            Logger.Log("In case a solution is open, something went wrong loading it.");
            Logger.Log("Maybe it helps to delete the solution .suo file and reload the solution?");
        }

        private bool _ConfigRefersToProject(MultiProjectConfigurationProject configProject, SolutionProject project)
        {
            return configProject.NameOrPath == project.Name ||
                   configProject.NameOrPath == project.Path;
        }
    }
}
