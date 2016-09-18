using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.VisualStudio;
using Microsoft.VisualStudio.Shell.Interop;

namespace LucidConcepts.SwitchStartupProject
{
    public class ConfigurationFileTracker : IVsFileChangeEvents
    {
        private readonly IVsFileChangeEx fileChangeService;
        private uint fileChangeCookie;
        private Action onConfigurationFileChanged;

        public ConfigurationFileTracker(string configurationFilename, IVsFileChangeEx fileChangeService, Action onConfigurationFileChanged)
        {
            this.fileChangeService = fileChangeService;
            this.onConfigurationFileChanged = onConfigurationFileChanged;

            fileChangeService.AdviseFileChange(
                configurationFilename,
                (uint)(_VSFILECHANGEFLAGS.VSFILECHG_Add | _VSFILECHANGEFLAGS.VSFILECHG_Del | _VSFILECHANGEFLAGS.VSFILECHG_Size | _VSFILECHANGEFLAGS.VSFILECHG_Time),
                this,
                out fileChangeCookie);
        }

        public void Stop()
        {
            fileChangeService.UnadviseFileChange(fileChangeCookie);
        }

        #region IVsFileChangeEvents Members

        public int DirectoryChanged(string pszDirectory)
        {
            return VSConstants.S_OK;
        }

        public int FilesChanged(uint cChanges, string[] rgpszFile, uint[] rggrfChange)
        {
            // Don't need to check the arguments since we ever only track the settings file
            onConfigurationFileChanged();
            return VSConstants.S_OK;
        }

        #endregion
    }
}
