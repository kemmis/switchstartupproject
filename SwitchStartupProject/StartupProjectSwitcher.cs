using System;
using System.Collections.Generic;
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
        private const string mostRecentlyUsedListKey = "MRU";
        private const string multiProjectConfigurationsKey = "MultiProjectConfigurations";

        private readonly ConfigurationsPersister settingsPersister;
        private readonly OptionPage options;
        private readonly OleMenuCommand menuSwitchStartupProjectComboCommand;
        private readonly Action openOptionsPage;
        private readonly SwitchStartupProjectPackage.ActivityLogger logger;

        private readonly Dictionary<IVsHierarchy, string> proj2name = new Dictionary<IVsHierarchy, string>();
        private readonly Dictionary<string, string> name2projectPath = new Dictionary<string, string>();

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



        public StartupProjectSwitcher(OleMenuCommand combobox, OptionPage options, DTE dte, Package package, int mruCount, SwitchStartupProjectPackage.ActivityLogger logger)
        {
            logger.LogInfo("Entering constructor of StartupProjectSwitcher");
            this.menuSwitchStartupProjectComboCommand = combobox;
            this.options = options;
            this.dte = dte;
            this.openOptionsPage = () => package.ShowOptionPage(typeof(OptionPage));
            this.logger = logger;

            // initialize MRU list
            mruStartupProjects = new MRUList<string>(mruCount);

            // initialize type list
            typeStartupProjects = new List<string>();

            // initialize all list
            allStartupProjects = new List<string>();

            multiProjectConfigurations = new List<MultiProjectConfiguration>();

            settingsPersister = new ConfigurationsPersister(dte, ".startup.suo");
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
                    _ChangeMultiProjectConfigurationsInOptions();
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

        public void UpdateStartupProject(IVsHierarchy startupProject)
        {
            // When startup project is set through dropdown, don't do anything
            if (!reactToChangedEvent) return;

            // When startup project is set in solution explorer, update combobox
            if (null != startupProject && proj2name.ContainsKey(startupProject))
            {
                var newStartupProjectName = proj2name[startupProject];
                logger.LogInfo("New startup project was activated outside of combobox: {0}", newStartupProjectName);
                mruStartupProjects.Touch(newStartupProjectName);
                _PopulateStartupProjects();
                currentStartupProject = newStartupProjectName;
            }
            else
            {
                logger.LogInfo("New unknown startup project was activated outside of combobox");
                currentStartupProject = sentinel;
            }
            
        }

        public void BeforeOpenSolution(string solutionFileName)
        {
            logger.LogInfo("Starting to open solution: {0}", solutionFileName);
            openingSolution = true;
        }

        public void AfterOpenSolution()
        {
            openingSolution = false;
            logger.LogInfo("Finished to open solution");
            mruStartupProjects = new MRUList<string>(options.MostRecentlyUsedCount, settingsPersister.GetList(mostRecentlyUsedListKey));
            options.Configurations.Clear();
            logger.LogInfo("Loading multi project configuration for solution");
            settingsPersister.GetMultiProjectConfigurations(multiProjectConfigurationsKey).ForEach(options.Configurations.Add);
            // When solution is open: enable multi-project configuration
            logger.LogInfo("Enable multi project configuration");
            options.EnableMultiProjectConfiguration = true;
            _PopulateStartupProjects();
        }

        public void BeforeCloseSolution()
        {
            logger.LogInfo("Starting to close solution");
            logger.LogInfo("Storing solution specific configuration.");
            // When solution is about to be closed, store MRU list to settings
            settingsPersister.StoreList(mostRecentlyUsedListKey, mruStartupProjects);
            settingsPersister.StoreMultiProjectConfigurations(multiProjectConfigurationsKey, options.Configurations);
            settingsPersister.Persist();
        }

        public void AfterCloseSolution()
        {
            logger.LogInfo("Finished to close solution");
            // When solution is closed: choose no project
            currentStartupProject = sentinel;
            options.Configurations.Clear();
            logger.LogInfo("Disable multi project configuration");
            options.EnableMultiProjectConfiguration = false;
            _ClearProjects();
        }

        public void OpenProject(IVsHierarchy pHierarchy)
        {
            // When project is opened: register it and its name
            bool valid = true;
            object nameObj = null;
            object typeNameObj = null;
            object captionObj = null;
            valid &= pHierarchy.GetProperty((uint)VSConstants.VSITEMID.Root, (int)__VSHPROPID.VSHPROPID_Name, out nameObj) == VSConstants.S_OK;
            valid &= pHierarchy.GetProperty((uint)VSConstants.VSITEMID.Root, (int)__VSHPROPID.VSHPROPID_TypeName, out typeNameObj) == VSConstants.S_OK;
            valid &= pHierarchy.GetProperty((uint)VSConstants.VSITEMID.Root, (int)__VSHPROPID.VSHPROPID_Caption, out captionObj) == VSConstants.S_OK;
            if (valid)
            {
                var name = (string)nameObj;
                var typeName = (string)typeNameObj;

                logger.LogInfo("Opening project: {0}", name);

                _AddProject(name, pHierarchy);
                allStartupProjects.Add(name);

                // Only add local (not web) C# projects with OutputType of either Exe (command line tool) or WinExe (windows application)
                if (typeName == "Microsoft Visual C# 2010")
                {
                    var project = _GetProject(pHierarchy);
                    var projectType = project.Properties.Item("ProjectType");
                    var eProjectType = (VSLangProj.prjProjectType)projectType.Value;
                    if (eProjectType == VSLangProj.prjProjectType.prjProjectTypeLocal)
                    {
                        var outputType = project.Properties.Item("OutputType");
                        var eOutputType = (VSLangProj.prjOutputType)outputType.Value;
                        if (eOutputType == VSLangProj.prjOutputType.prjOutputTypeWinExe ||
                            eOutputType == VSLangProj.prjOutputType.prjOutputTypeExe)
                        {
                            typeStartupProjects.Add(name);
                        }
                    }
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
            if (startupProjects.Contains(projectName))
            {
                if (currentStartupProject == projectName)
                {
                    currentStartupProject = sentinel;
                }
                startupProjects.Remove(projectName);
                allStartupProjects.Remove(projectName);
                typeStartupProjects.Remove(projectName);
                proj2name.Remove(pHierarchy);
                name2projectPath.Remove(projectName);
            }
        }

        public void ToggleDebuggingActive(bool active)
        {
            logger.LogInfo(active ? "Start debugging, disable combobox" : "Stop debugging, enable combobox");
            // When debugging command UI context is activated, disable combobox, otherwise enable combobox
            menuSwitchStartupProjectComboCommand.Enabled = !active;
        }


        private void _ChangeStartupProject(string newStartupProject)
        {
            this.currentStartupProject = newStartupProject;
            if (newStartupProject == sentinel)
            {
                // No startup project
                _SuspendChangedEvent(() => dte.Solution.SolutionBuild.StartupProjects = null);
            }
            else if (name2projectPath.ContainsKey(newStartupProject))
            {
                // Single startup project
                _SuspendChangedEvent(() => dte.Solution.SolutionBuild.StartupProjects = name2projectPath[newStartupProject]);
            }
            else if (multiProjectConfigurations.Any(c => c.Name == newStartupProject))
            {
                // Multiple startup projects
                var configuration = multiProjectConfigurations.Single(c => c.Name == newStartupProject);
                _SuspendChangedEvent(() => dte.Solution.SolutionBuild.StartupProjects = configuration.Projects.OfType<object>().ToArray()); // SolutionBuild.StartupProjects expects an array of objects
            }
            // An unknown project was chosen
        }

        private void _SuspendChangedEvent(Action action)
        {
            reactToChangedEvent = false;
            action();
            reactToChangedEvent = true;
        }

        private Project _GetProject(IVsHierarchy pHierarchy)
        {
            object project;
            ErrorHandler.ThrowOnFailure(pHierarchy.GetProperty(VSConstants.VSITEMID_ROOT,
                                                               (int)__VSHPROPID.VSHPROPID_ExtObject,
                                                               out project));
            return (project as Project);
        }

        private string _GetPathRelativeToSolution(IVsHierarchy pHierarchy)
        {
            var project = _GetProject(pHierarchy);
            var fullProjectPath = project.FullName;
            var solutionPath = Path.GetDirectoryName(dte.Solution.FullName) + @"\";
            return Paths.GetPathRelativeTo(fullProjectPath, solutionPath);
        }

        private void _AddProject(string name, IVsHierarchy pHierarchy)
        {
            proj2name.Add(pHierarchy, name);
            name2projectPath.Add(name, _GetPathRelativeToSolution(pHierarchy));
        }

        private void _ClearProjects()
        {
            proj2name.Clear();
            name2projectPath.Clear();
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

            if (currentStartupProject != sentinel && !startupProjects.Contains(currentStartupProject))
            {
                _ChangeStartupProject(sentinel);
            }
        }

        private void _ChangeMultiProjectConfigurationsInOptions()
        {
            multiProjectConfigurations.Clear();
            multiProjectConfigurations = (from configuration in options.Configurations
                                          let projects = (from project in configuration.Projects
                                                          where project.Name != null && name2projectPath.ContainsKey(project.Name)
                                                          select name2projectPath[project.Name]).ToList()
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
