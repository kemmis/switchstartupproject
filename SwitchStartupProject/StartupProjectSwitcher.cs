using System;
using System.Collections.Generic;
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

        private readonly ConfigurationsPersister settingsPersister;
        private readonly OptionPage options;
        private readonly OleMenuCommand menuSwitchStartupProjectComboCommand;
        private readonly Action openOptionsPage;


        private readonly Dictionary<IVsHierarchy, string> proj2name = new Dictionary<IVsHierarchy, string>();
        private readonly Dictionary<string, IVsHierarchy> name2proj = new Dictionary<string, IVsHierarchy>();

        private MRUList<string> mruStartupProjects;
        private List<string> typeStartupProjects;
        private List<string> allStartupProjects;

        private const string sentinel = "";
        private const string configure = "Configure...";
        private List<string> startupProjects = new List<string>(new [] { configure });
        private string currentStartupProject = sentinel;
        private bool openingSolution;

        private readonly IVsSolutionBuildManager2 sbm;



        public StartupProjectSwitcher(OleMenuCommand combobox, OptionPage options, DTE dte, IVsSolutionBuildManager2 sbm, Package package, int mruCount)
        {
            this.menuSwitchStartupProjectComboCommand = combobox;
            this.options = options;
            this.sbm = sbm;
            this.openOptionsPage = () => package.ShowOptionPage(typeof (OptionPage));


            // initialize MRU list
            mruStartupProjects = new MRUList<string>(mruCount);

            // initialize type list
            typeStartupProjects = new List<string>();

            // initialize all list
            allStartupProjects = new List<string>();

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
            };
            
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
                openOptionsPage();
                return;

            }

            // see if it is one of our items
            foreach (string project in this.startupProjects)
            {
                if (String.Compare(project, name, StringComparison.CurrentCultureIgnoreCase) == 0)
                {
                    _ChangeStartupProject(project);
                    return;
                }
            }
            throw (new ArgumentException("ParamNotValidStringInList")); // force an exception to be thrown
        }

        public void UpdateStartupProject(IVsHierarchy startupProject)
        {
            // When startup project is set in solution explorer, update combobox
            if (null != startupProject && proj2name.ContainsKey(startupProject))
            {
                var newStartupProjectName = proj2name[startupProject];
                mruStartupProjects.Touch(newStartupProjectName);
                _PopulateStartupProjects();
                currentStartupProject = newStartupProjectName;
            }
            else
            {
                currentStartupProject = sentinel;
            }
            
        }

        public void BeforeOpenSolution(string solutionFileName)
        {
            openingSolution = true;
        }

        public void AfterOpenSolution()
        {
            openingSolution = false;
            mruStartupProjects = new MRUList<string>(options.MostRecentlyUsedCount, settingsPersister.GetList(mostRecentlyUsedListKey));
            // When solution is open: enable combobox
            menuSwitchStartupProjectComboCommand.Enabled = true;
            _PopulateStartupProjects();
        }

        public void BeforeCloseSolution()
        {
            // When solution is about to be closed, store MRU list to settings
            settingsPersister.StoreList(mostRecentlyUsedListKey, mruStartupProjects);
            settingsPersister.Persist();
        }

        public void AfterCloseSolution()
        {
            // When solution is closed: choose no project, disable combobox
            currentStartupProject = sentinel;
            menuSwitchStartupProjectComboCommand.Enabled = false;
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

                _AddProject(name, pHierarchy);
                allStartupProjects.Add(name);

                // Only add local (not web) C# projects with OutputType of either Exe (command line tool) or WinExe (windows application)
                if (typeName == "Microsoft Visual C# 2010")
                {
                    var project = GetProject(pHierarchy);
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
                name2proj.Remove(projectName);
                proj2name.Remove(pHierarchy);
            }
        }

        public void ToggleDebuggingActive(bool active)
        {
            // When debugging command UI context is activated, disable combobox, otherwise enable combobox
            menuSwitchStartupProjectComboCommand.Enabled = !active;
        }


        private void _ChangeStartupProject(string newStartupProject)
        {
            this.currentStartupProject = newStartupProject;
            if (!name2proj.ContainsKey(newStartupProject)) return;  // An unknown project was chosen
            if (newStartupProject == sentinel) return;  // Sentinel was chosen
            IVsHierarchy oldStartupProject = null;
            if (sbm.get_StartupProject(out oldStartupProject) != VSConstants.S_OK) return;  // Can't get old startup project
            if (proj2name.ContainsKey(oldStartupProject) && newStartupProject == proj2name[oldStartupProject]) return;  // The chosen project was already the startup project
            sbm.set_StartupProject(name2proj[newStartupProject]);
        }

        private Project GetProject(IVsHierarchy pHierarchy)
        {
            object project;
            ErrorHandler.ThrowOnFailure(pHierarchy.GetProperty(VSConstants.VSITEMID_ROOT,
                                                               (int)__VSHPROPID.VSHPROPID_ExtObject,
                                                               out project));
            return (project as Project);
        }

        private void _AddProject(string name, IVsHierarchy pHierarchy)
        {
            name2proj.Add(name, pHierarchy);
            proj2name.Add(pHierarchy, name);
        }

        private void _ClearProjects()
        {
            name2proj.Clear();
            proj2name.Clear();
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
            }
            startupProjects.Add(configure);
        }

        private void _ChangeMRUCountInOptions()
        {
            var oldList = mruStartupProjects;
            mruStartupProjects = new MRUList<string>(options.MostRecentlyUsedCount, oldList);
            _PopulateStartupProjects();
        }
    }
}
