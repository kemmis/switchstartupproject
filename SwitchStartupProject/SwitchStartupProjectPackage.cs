using System;
using System.Collections.Generic;
using System.ComponentModel.Design;
using System.Globalization;
using System.Linq;
using System.Runtime.InteropServices;

using Microsoft.VisualStudio;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;

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
    public sealed class SwitchStartupProjectPackage : Package, IVsSolutionEvents, IVsSolutionLoadEvents, IVsSelectionEvents, IVsPersistSolutionOpts
    {

        private uint solutionEventsCookie;
        private IVsSolution2 solution = null;
        private uint selectionEventsCookie;
        private IVsMonitorSelection ms = null;
        private uint debuggingCookie;
        private bool projectsAreLoadedInBatches = false;

        private StartupProjectSwitcher switcher;

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
            Logger = new ActivityLogger(this);
            Logger.LogInfo("Entering initializer for: {0}", this.ToString());
            base.Initialize();

            OleMenuCommand menuSwitchStartupProjectComboCommand = null;

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
                menuSwitchStartupProjectComboCommand = new OleMenuCommand(new EventHandler(_OnMenuSwitchStartupProjectCombo), menuSwitchStartupProjectComboCommandID);
                menuSwitchStartupProjectComboCommand.ParametersDescription = "$"; // accept any argument string
                mcs.AddCommand(menuSwitchStartupProjectComboCommand);
                menuSwitchStartupProjectComboCommand.Enabled = false;


                CommandID menuSwitchStartupProjectComboGetListCommandID = new CommandID(GuidList.guidSwitchStartupProjectCmdSet, (int)PkgCmdIDList.cmdidSwitchStartupProjectComboGetList);
                MenuCommand menuSwitchStartupProjectComboGetListCommand = new OleMenuCommand(new EventHandler(_OnMenuSwitchStartupProjectComboGetList), menuSwitchStartupProjectComboGetListCommandID);
                mcs.AddCommand(menuSwitchStartupProjectComboGetListCommand);
            }

            // Get VS Automation object
            var dte = (EnvDTE.DTE)GetGlobalService(typeof(EnvDTE.DTE));

            // Get solution
            solution = ServiceProvider.GlobalProvider.GetService(typeof(SVsSolution)) as IVsSolution2;
            if (solution != null)
            {
                // Register for solution events
                solution.AdviseSolutionEvents(this, out solutionEventsCookie);
            }

            // get options
            var options = (OptionPage)GetDialogPage(typeof(OptionPage));
            options.Logger = Logger;

            // Get selection monitor
            ms = ServiceProvider.GlobalProvider.GetService(typeof(SVsShellMonitorSelection)) as IVsMonitorSelection;
            if (ms != null)
            {
                // Remember debugging UI context cookie for later
                ms.GetCmdUIContextCookie(VSConstants.UICONTEXT.Debugging_guid, out debuggingCookie);
                // Register for selection events
                ms.AdviseSelectionEvents(this, out selectionEventsCookie);
            }

            var fileChangeService = ServiceProvider.GlobalProvider.GetService(typeof(SVsFileChangeEx)) as IVsFileChangeEx;

            switcher = new StartupProjectSwitcher(menuSwitchStartupProjectComboCommand, options, dte, fileChangeService, this, options.MostRecentlyUsedCount, Logger);
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
        private void _OnMenuSwitchStartupProjectCombo(object sender, EventArgs e)
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
                    Marshal.GetNativeVariantForObject(switcher.GetCurrentStartupProject(), vOut);
                }

                else if (newChoice != null)
                {
                    switcher.ChooseStartupProject(newChoice);
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
        private void _OnMenuSwitchStartupProjectComboGetList(object sender, EventArgs e)
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
                    Marshal.GetNativeVariantForObject(switcher.GetStartupProjectChoices(), vOut);
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
            switcher.AfterCloseSolution();
            return VSConstants.S_OK;
        }

        public int OnAfterLoadProject(IVsHierarchy pStubHierarchy, IVsHierarchy pRealHierarchy)
        {
            return VSConstants.S_OK;
        }

        public int OnAfterOpenProject(IVsHierarchy pHierarchy, int fAdded)
        {
            switcher.OpenProject(pHierarchy);
            return VSConstants.S_OK;
        }


        public int OnAfterOpenSolution(object pUnkReserved, int fNewSolution)
        {
            if (!projectsAreLoadedInBatches)
            {
                switcher.AfterOpenSolution();
            }
            return VSConstants.S_OK;
        }

        public int OnBeforeCloseProject(IVsHierarchy pHierarchy, int fRemoved)
        {
            object propNameObj = null;
            if (pHierarchy.GetProperty((uint)VSConstants.VSITEMID.Root, (int)__VSHPROPID.VSHPROPID_Name, out propNameObj) == VSConstants.S_OK)
            {
                string name = (string)propNameObj;
                switcher.CloseProject(pHierarchy, name);
            }
            return VSConstants.S_OK;
        }

        public int OnBeforeCloseSolution(object pUnkReserved)
        {
            switcher.BeforeCloseSolution();
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
            switcher.AfterOpenSolution();
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
            projectsAreLoadedInBatches = true;
            return VSConstants.S_OK;
        }

        public int OnBeforeOpenSolution(string pszSolutionFilename)
        {
            projectsAreLoadedInBatches = false;
            switcher.BeforeOpenSolution(pszSolutionFilename);
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
            if (dwCmdUICookie == debuggingCookie)
            {
                switcher.ToggleDebuggingActive(fActive != 0);
            }
            return VSConstants.S_OK;
        }

        public int OnElementValueChanged(uint elementid, object varValueOld, object varValueNew)
        {
            if (elementid == (uint)VSConstants.VSSELELEMID.SEID_StartupProject)
            {
                switcher.UpdateStartupProject();
            }
            return VSConstants.S_OK;
        }

        public int OnSelectionChanged(IVsHierarchy pHierOld, uint itemidOld, IVsMultiItemSelect pMISOld, ISelectionContainer pSCOld, IVsHierarchy pHierNew, uint itemidNew, IVsMultiItemSelect pMISNew, ISelectionContainer pSCNew)
        {
            return VSConstants.S_OK;
        }

        #endregion

        #region IVsPersistSolutionOpts Members

        public int SaveUserOptions(IVsSolutionPersistence pPersistence)
        {
            switcher.OnSolutionSaved();
            return VSConstants.S_OK;
        }

        #endregion

        #region Activity Log

        public ActivityLogger Logger { get; set; }

        public class ActivityLogger
        {
            private readonly SwitchStartupProjectPackage package;

            public ActivityLogger(SwitchStartupProjectPackage package)
            {
                this.package = package;
            }

            public void LogInfo(string message, params object[] arguments)
            {
                Log(__ACTIVITYLOG_ENTRYTYPE.ALE_INFORMATION, message, arguments);
            }

            public void LogWarning(string message, params object[] arguments)
            {
                Log(__ACTIVITYLOG_ENTRYTYPE.ALE_WARNING, message, arguments);
            }

            public void LogError(string message, params object[] arguments)
            {
                Log(__ACTIVITYLOG_ENTRYTYPE.ALE_ERROR, message, arguments);
            }

            private void Log(__ACTIVITYLOG_ENTRYTYPE type, string message, params object[] arguments)
            {
                var log = package.GetService(typeof (SVsActivityLog)) as IVsActivityLog;
                if (log == null) return;
                log.LogEntry((UInt32)type, "SwitchStartupProject", string.Format(CultureInfo.CurrentCulture, message, arguments));
            }
        }

        #endregion
    }
}
