using System;
using System.Collections.Generic;
using System.ComponentModel.Design;
using System.Diagnostics;
using System.Globalization;
using System.Linq;
using System.Runtime.InteropServices;

using Microsoft.VisualStudio;
using Microsoft.VisualStudio.OLE.Interop;
using Microsoft.VisualStudio.Package;
using Microsoft.VisualStudio.Settings;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using Microsoft.Win32;

using EnvDTE;
using Microsoft.VisualStudio.Shell.Settings;

namespace LucidConcepts.SwitchStartupProject
{
    /// <summary>
    /// This is the class that implements the package exposed by this assembly.
    ///
    /// The minimum requirement for a class to be considered a valid package for Visual Studio
    /// is to implement the IVsPackage interface and register itself with the shell.
    /// This package uses the helper classes defined inside the Managed Package Framework (MPF)
    /// to do it: it derives from the Package class that provides the implementation of the 
    /// IVsPackage interface and uses the registration attributes defined in the framework to 
    /// register itself and its components with the shell.
    /// </summary>
    // This attribute tells the PkgDef creation utility (CreatePkgDef.exe) that this class is
    // a package.
    [PackageRegistration(UseManagedResourcesOnly = true)]
    // This attribute is used to register the informations needed to show the this package
    // in the Help/About dialog of Visual Studio.
    [InstalledProductRegistration("#110", "#112", "1.0", IconResourceID = 400)]
    // This attribute is needed to let the shell know that this package exposes some menus.
    [ProvideMenuResource("Menus.ctmenu", 1)]
    // This attribute says that this package exposes an options page of the given type.
    [ProvideOptionPage(typeof(OptionPage), "Switch Startup Project", "General", 0, 0, true)]
    [Guid(GuidList.guidSwitchStartupProjectPkgString)]
    public sealed class SwitchStartupProjectPackage : Package, IVsSolutionEvents, IVsSolutionLoadEvents, IVsSelectionEvents
    {
        OleMenuCommand menuSwitchStartupProjectComboCommand;

        private uint solutionEventsCookie;
        private IVsSolution2 solution = null;
        private IVsSolutionBuildManager2 sbm2 = null;
        private uint selectionEventsCookie;
        private IVsMonitorSelection ms = null;
        private uint debuggingCookie;

        private MRUList<string> mruStartupProjects;
        private List<string> typeStartupProjects;

        private Dictionary<IVsHierarchy, string> proj2name = new Dictionary<IVsHierarchy, string>();
        private Dictionary<string, IVsHierarchy> name2proj = new Dictionary<string, IVsHierarchy>();
        private const string sentinel = "";
        private List<string> startupProjects = new List<string>(new string[] { sentinel });
        private string currentStartupProject = sentinel;
        private string currentSolutionFilename;
        private OptionPage options;
        private WritableSettingsStore userSettingsStore;

        protected override void Dispose(bool disposing)
        {
            base.Dispose(disposing);

            // Unadvise all events
            if (solution != null && solutionEventsCookie != 0)
                solution.UnadviseSolutionEvents(solutionEventsCookie);
            if (ms != null && selectionEventsCookie != 0)
                ms.UnadviseSelectionEvents(selectionEventsCookie);
        }

        #region Package Initializer

        /// <summary>
        /// Initialization of the package; this method is called right after the package is sited, so this is the place
        /// where you can put all the initilaization code that rely on services provided by VisualStudio.
        /// </summary>
        protected override void Initialize()
        {
            LogStart();
            base.Initialize();

            // get options
            options = (OptionPage)GetDialogPage(typeof(OptionPage));
            options.Modified += (s, e) =>
            {
                if (e.OptionParameter == EOptionParameter.MruMode)
                {
                    _SwitchModeInOptions();
                }
                else if (e.OptionParameter == EOptionParameter.MruCount)
                {
                    _ChangeMRUCountInOptions();
                }
            };

            // get settings store
            SettingsManager settingsManager = new ShellSettingsManager(this);
            userSettingsStore = settingsManager.GetWritableSettingsStore(SettingsScope.UserSettings);

            // Add our command handlers for menu (commands must exist in the .vsct file)
            OleMenuCommandService mcs = GetService(typeof(IMenuCommandService)) as OleMenuCommandService;
            if ( null != mcs )
            {
                // DropDownCombo
                //	 a DROPDOWNCOMBO does not let the user type into the combo box; they can only pick from the list.
                //   The string value of the element selected is returned.
                //	 For example, this type of combo could be used for the "Solution Configurations" on the "Standard" toolbar.
                //
                //   A DropDownCombo box requires two commands:
                //     One command (cmdidMyCombo) is used to ask for the current value of the combo box and to 
                //     set the new value when the user makes a choice in the combo box.
                //
                //     The second command (cmdidMyComboGetList) is used to retrieve this list of choices for the combo box.
                CommandID menuSwitchStartupProjectComboCommandID = new CommandID(GuidList.guidSwitchStartupProjectCmdSet, (int)PkgCmdIDList.cmdidSwitchStartupProjectCombo);
                menuSwitchStartupProjectComboCommand = new OleMenuCommand(new EventHandler(OnMenuSwitchStartupProjectCombo), menuSwitchStartupProjectComboCommandID);
                menuSwitchStartupProjectComboCommand.ParametersDescription = "$"; // accept any argument string
                mcs.AddCommand(menuSwitchStartupProjectComboCommand);
                menuSwitchStartupProjectComboCommand.Enabled = false;


                CommandID menuSwitchStartupProjectComboGetListCommandID = new CommandID(GuidList.guidSwitchStartupProjectCmdSet, (int)PkgCmdIDList.cmdidSwitchStartupProjectComboGetList);
                MenuCommand menuSwitchStartupProjectComboGetListCommand = new OleMenuCommand(new EventHandler(OnMenuSwitchStartupProjectComboGetList), menuSwitchStartupProjectComboGetListCommandID);
                mcs.AddCommand(menuSwitchStartupProjectComboGetListCommand);
            }

            // Get solution
            solution = ServiceProvider.GlobalProvider.GetService(typeof(SVsSolution)) as IVsSolution2;
            if (solution != null)
            {
                // Register for solution events
                solution.AdviseSolutionEvents(this, out solutionEventsCookie);
            }

            // Get solution build manager
            sbm2 = ServiceProvider.GlobalProvider.GetService(typeof(SVsSolutionBuildManager)) as IVsSolutionBuildManager2;

            // Get selection monitor
            ms = ServiceProvider.GlobalProvider.GetService(typeof(SVsShellMonitorSelection)) as IVsMonitorSelection;
            if (ms != null)
            {
                // Remember debugging UI context cookie for later
                ms.GetCmdUIContextCookie(VSConstants.UICONTEXT.Debugging_guid, out debuggingCookie);
                // Register for selection events
                ms.AdviseSelectionEvents(this, out selectionEventsCookie);
            }

            // initialize MRU list
            mruStartupProjects = new MRUList<string>(options.MruCount);

            // initialize type list
            typeStartupProjects = new List<string>();

        }
        #endregion

        #region ComboBoxHandlers

        // IndexCombo
        //	 An INDEXCOMBO is the same as a DROPDOWNCOMBO in that it is a "pick from list" only combo.
        //	 The difference is an INDEXCOMBO returns the selected value as an index into the list (0 based).
        //	 For example, this type of combo could be used for the "Solution Configurations" on the "Standard" toolbar.
        //
        //   An IndexCombo box requires two commands:
        //     One command is used to ask for the current value of the combo box and to set the new value when the user
        //     makes a choice in the combo box.
        //
        //     The second command is used to retrieve this list of choices for the combo box.
        private void OnMenuSwitchStartupProjectCombo(object sender, EventArgs e)
        {
            if ((null == e) || (e == EventArgs.Empty))
            {
                // We should never get here; EventArgs are required.
                throw (new ArgumentException("EventArgs are required")); // force an exception to be thrown
            }

            OleMenuCmdEventArgs eventArgs = e as OleMenuCmdEventArgs;

            if (eventArgs != null)
            {
                string newChoice = eventArgs.InValue as string;
                IntPtr vOut = eventArgs.OutValue;

                if (vOut != IntPtr.Zero && newChoice != null)
                {
                    throw (new ArgumentException("Both in and out parameters should not be specified")); // force an exception to be thrown
                }
                if (vOut != IntPtr.Zero)
                {
                    // when vOut is non-NULL, the IDE is requesting the current value for the combo
                    Marshal.GetNativeVariantForObject(this.currentStartupProject, vOut);
                }

                else if (newChoice != null)
                {
                    // new value was selected or typed in
                    // see if it is one of our items
                    bool validInput = false;
                    int indexInput = -1;
                    for (indexInput = 0; indexInput < startupProjects.Count; indexInput++)
                    {
                        if (String.Compare(startupProjects[indexInput], newChoice, StringComparison.CurrentCultureIgnoreCase) == 0)
                        {
                            validInput = true;
                            break;
                        }
                    }

                    if (validInput)
                    {
                        _SetStartupProjectInCombo(startupProjects[indexInput]);
                    }
                    else
                    {
                        throw (new ArgumentException("ParamNotValidStringInList")); // force an exception to be thrown
                    }
                }
                else
                {
                    // We should never get here
                    throw (new ArgumentException("InOutParamCantBeNULL")); // force an exception to be thrown
                }
            }
            else
            {
                // We should never get here; EventArgs are required.
                throw (new ArgumentException("EventArgs are required")); // force an exception to be thrown
            }
        }

        // An IndexCombo box requires two commands:
        //    This command is used to retrieve this list of choices for the combo box.
        // 
        // Normally IOleCommandTarget::QueryStatus is used to determine the state of a command, e.g.
        // enable vs. disable, shown vs. hidden, etc. The QueryStatus method does not have any way to 
        // control the statue of a combo box, e.g. what list of items should be shown and what is the 
        // current value. In order to communicate this information actually IOleCommandTarget::Exec
        // is used with a non-NULL varOut parameter. You can think of these Exec calls as extended 
        // QueryStatus calls. There are two pieces of information needed for a combo, thus it takes
        // two commands to retrieve this information. The main command id for the command is used to 
        // retrieve the current value and the second command is used to retrieve the full list of 
        // choices to be displayed as an array of strings.
        private void OnMenuSwitchStartupProjectComboGetList(object sender, EventArgs e)
        {
            if (e == EventArgs.Empty)
            {
                // We should never get here; EventArgs are required.
                throw (new ArgumentException("EventArgs are required")); // force an exception to be thrown
            }

            OleMenuCmdEventArgs eventArgs = e as OleMenuCmdEventArgs;

            if (eventArgs != null)
            {
                object inParam = eventArgs.InValue;
                IntPtr vOut = eventArgs.OutValue;

                if (inParam != null)
                {
                    throw (new ArgumentException("In parameter may not be specified")); // force an exception to be thrown
                }
                else if (vOut != IntPtr.Zero)
                {
                    Marshal.GetNativeVariantForObject(this.startupProjects.ToArray(), vOut);
                }
                else
                {
                    throw (new ArgumentException("Out parameter can not be NULL")); // force an exception to be thrown
                }
            }
        }


        #endregion

        #region IVsSolutionEvents Members

        public int OnAfterCloseSolution(object pUnkReserved)
        {
            // When solution is closed: choose no project, disable combobox
            currentStartupProject = sentinel;
            menuSwitchStartupProjectComboCommand.Enabled = false;
            _ClearProjects();
            return VSConstants.S_OK;
        }

        public int OnAfterLoadProject(IVsHierarchy pStubHierarchy, IVsHierarchy pRealHierarchy)
        {
            return VSConstants.S_OK;
        }

        public int OnAfterOpenProject(IVsHierarchy pHierarchy, int fAdded)
        {
            _RegisterProject(pHierarchy);
            return VSConstants.S_OK;
        }


        public int OnAfterOpenSolution(object pUnkReserved, int fNewSolution)
        {
            // When solution is open: enable combobox
            menuSwitchStartupProjectComboCommand.Enabled = true;
            if (!options.MruMode)
            {
                _PopulateStartupProjectsFromTypeList();
            }
            return VSConstants.S_OK;
        }

        public int OnBeforeCloseProject(IVsHierarchy pHierarchy, int fRemoved)
        {
            // When project is closed: remove it from list of startup projects (if it was in there)
            object propNameObj = null;
            if (pHierarchy.GetProperty((uint)VSConstants.VSITEMID.Root, (int)__VSHPROPID.VSHPROPID_Name, out propNameObj) == VSConstants.S_OK)
            {
                string name = (string)propNameObj;
                if (startupProjects.Contains(name))
                {
                    if (currentStartupProject == name)
                    {
                        currentStartupProject = sentinel;
                    }
                    startupProjects.Remove(name);
                    name2proj.Remove(name);
                    proj2name.Remove(pHierarchy);

                }
            }
            return VSConstants.S_OK;
        }

        public int OnBeforeCloseSolution(object pUnkReserved)
        {
            // When solution is about to be closed, store MRU list to settings
            _StoreMRUListToSettings();
            return VSConstants.S_OK;
        }

        public int OnBeforeUnloadProject(IVsHierarchy pRealHierarchy, IVsHierarchy pStubHierarchy)
        {
            return VSConstants.S_OK;
        }

        public int OnQueryCloseProject(IVsHierarchy pHierarchy, int fRemoving, ref int pfCancel)
        {
            return VSConstants.S_OK;
        }

        public int OnQueryCloseSolution(object pUnkReserved, ref int pfCancel)
        {
            return VSConstants.S_OK;
        }

        public int OnQueryUnloadProject(IVsHierarchy pRealHierarchy, ref int pfCancel)
        {
            return VSConstants.S_OK;
        }

        #endregion

        #region IVsSolutionLoadEvents Members

        public int OnAfterBackgroundSolutionLoadComplete()
        {
            return VSConstants.S_OK;
        }

        public int OnAfterLoadProjectBatch(bool fIsBackgroundIdleBatch)
        {
            return VSConstants.S_OK;
        }

        public int OnBeforeBackgroundSolutionLoadBegins()
        {
            return VSConstants.S_OK;
        }

        public int OnBeforeLoadProjectBatch(bool fIsBackgroundIdleBatch)
        {
            return VSConstants.S_OK;
        }

        public int OnBeforeOpenSolution(string pszSolutionFilename)
        {
            currentSolutionFilename = pszSolutionFilename;
            _PopulateMRUListFromSettings(pszSolutionFilename);
            return VSConstants.S_OK;
        }

        public int OnQueryBackgroundLoadProjectBatch(out bool pfShouldDelayLoadToNextIdle)
        {
            pfShouldDelayLoadToNextIdle = false;
            return VSConstants.S_OK;
        }

        #endregion

        #region IVsSelectionEvents Members

        public int OnCmdUIContextChanged(uint dwCmdUICookie, int fActive)
        {
            // When debugging command UI context is activated, disable combobox, otherwise enable combobox
            if (dwCmdUICookie == debuggingCookie)
            {
                menuSwitchStartupProjectComboCommand.Enabled = fActive == 0;
            }
            return VSConstants.S_OK;
        }

        public int OnElementValueChanged(uint elementid, object varValueOld, object varValueNew)
        {
            if (elementid == (uint)VSConstants.VSSELELEMID.SEID_StartupProject)
            {
                // When startup project is set in solution explorer, update combobox
                _SetStartupProjectInTree((IVsHierarchy)varValueNew);
            }
            return VSConstants.S_OK;
        }

        public int OnSelectionChanged(IVsHierarchy pHierOld, uint itemidOld, IVsMultiItemSelect pMISOld, ISelectionContainer pSCOld, IVsHierarchy pHierNew, uint itemidNew, IVsMultiItemSelect pMISNew, ISelectionContainer pSCNew)
        {
            return VSConstants.S_OK;
        }

        #endregion

        #region project management

        private void _RegisterProject(IVsHierarchy pHierarchy)
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
                string name = (string)nameObj;
                string typeName = (string)typeNameObj;
                string caption = (string)captionObj;

                _AddProject(name, pHierarchy);

                // Only add local (not web) C# projects with OutputType of either Exe (command line tool) or WinExe (windows application)
                if (typeName == "Microsoft Visual C# 2010")
                {
                    Project project = GetProject(pHierarchy);
                    var projectType = project.Properties.Item("ProjectType");
                    var eProjectType = (VSLangProj.prjProjectType)projectType.Value;
                    if (eProjectType == VSLangProj.prjProjectType.prjProjectTypeLocal)
                    {
                        var outputType = project.Properties.Item("OutputType");
                        var eOutputType = (VSLangProj.prjOutputType)outputType.Value;
                        if (eOutputType == VSLangProj.prjOutputType.prjOutputTypeWinExe ||
                            eOutputType == VSLangProj.prjOutputType.prjOutputTypeExe)
                        {
                            _AddStartupProjectType(name, pHierarchy);
                        }
                    }
                }
            }
        }

        private void _ClearProjects()
        {
            name2proj.Clear();
            proj2name.Clear();
            typeStartupProjects = new List<string>();
            startupProjects = new List<string>(new string[] { sentinel });
        }

        private void _AddProject(string name, IVsHierarchy pHierarchy)
        {
            name2proj.Add(name, pHierarchy);
            proj2name.Add(pHierarchy, name);
        }

        private void _AddStartupProjectType(string name, IVsHierarchy pHierarchy)
        {
            typeStartupProjects.Add(name);
        }

        private void _SetStartupProjectInCombo(string newStartupProject)
        {
            IVsHierarchy oldStartupProject = null;
            if (sbm2.get_StartupProject(out oldStartupProject) == VSConstants.S_OK)
            {
                if ((!proj2name.ContainsKey(oldStartupProject) || newStartupProject != proj2name[oldStartupProject]) &&
                    name2proj.ContainsKey(newStartupProject) &&
                    newStartupProject != sentinel)
                {
                    sbm2.set_StartupProject(name2proj[newStartupProject]);
                }
            }
            this.currentStartupProject = newStartupProject;
        }

        private void _SetStartupProjectInTree(IVsHierarchy startupProject)
        {
            if (null != startupProject && proj2name.ContainsKey(startupProject))
            {
                var newStartupProjectName = proj2name[startupProject];
                mruStartupProjects.Touch(newStartupProjectName);
                if (options.MruMode)
                {
                    _PopulateStartupProjectsFromMRUList();
                }
                currentStartupProject = newStartupProjectName;
            }
            else
            {
                currentStartupProject = sentinel;
            }
        }

        private void _PopulateStartupProjectsFromTypeList()
        {
            typeStartupProjects.Sort();
            startupProjects = typeStartupProjects.ToList();
        }

        private void _PopulateStartupProjectsFromMRUList()
        {
            startupProjects = mruStartupProjects.ToList();
        }

        private const string collectionKeyFormat = "LucidConcepts\\SwitchStartupProject\\StartupProjectMRUListBySolution\\{0}";
        private void _PopulateMRUListFromSettings(string solutionFileName)
        {
            var collectionPath = string.Format(collectionKeyFormat, _GetPathFromSolutionFilename(solutionFileName));
            if (userSettingsStore.CollectionExists(collectionPath))
            {
                var names = userSettingsStore.GetPropertyNames(collectionPath).ToList();
                names.Sort();
                mruStartupProjects = new MRUList<string>(options.MruCount, names.Select(name => userSettingsStore.GetString(collectionPath, name)));
            }
            else
            {
                mruStartupProjects = new MRUList<string>(options.MruCount);
            }
        }

        private void _StoreMRUListToSettings()
        {
            var collectionPath = string.Format(collectionKeyFormat, _GetPathFromSolutionFilename(currentSolutionFilename));
            if (userSettingsStore.CollectionExists(collectionPath))
            {
                userSettingsStore.DeleteCollection(collectionPath);
            }
            userSettingsStore.CreateCollection(collectionPath);
            int index = 0;
            foreach (var project in mruStartupProjects)
            {
                userSettingsStore.SetString(collectionPath, string.Format("Project{0:D2}", index++), project);
            }
        }

        private string _GetPathFromSolutionFilename(string filename)
        {
            return filename.Replace('\\', '/');
        }

        private void _SwitchModeInOptions()
        {
            if (options.MruMode)
            {
                _PopulateStartupProjectsFromMRUList();
            }
            else
            {
                _PopulateStartupProjectsFromTypeList();
            }
        }

        private void _ChangeMRUCountInOptions()
        {
            var oldList = mruStartupProjects;
            mruStartupProjects = new MRUList<string>(options.MruCount, oldList);
            if (options.MruMode)
            {
                _PopulateStartupProjectsFromMRUList();
            }
        }

        #endregion

        #region helper methods

        private Project GetProject(IVsHierarchy pHierarchy)
        {
            object project;
            ErrorHandler.ThrowOnFailure(pHierarchy.GetProperty(VSConstants.VSITEMID_ROOT,
                                                               (int)__VSHPROPID.VSHPROPID_ExtObject,
                                                               out project));
            return (project as Project);
        }

        private void LogStart()
        {
            IVsActivityLog log = GetService(typeof(SVsActivityLog)) as IVsActivityLog;
            if (log == null) return;
            var name = this.ToString();
            int hr = log.LogEntry((UInt32)__ACTIVITYLOG_ENTRYTYPE.ALE_INFORMATION,
                name,
                string.Format(CultureInfo.CurrentCulture,
                "Entering initializer for: {0}", name));
        }

        #endregion

    }
}
