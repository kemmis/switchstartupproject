using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Runtime.InteropServices;
using Microsoft.VisualStudio.Shell;
using System.ComponentModel;

namespace LucidConcepts.SwitchStartupProject
{
    public enum EOptionParameter
    {
        Mode,
        MruCount
    }

    public enum EMode
    {
        Smart,
        Mru
    }

    public class OptionsModifiedEventArgs : EventArgs
    {
        public OptionsModifiedEventArgs(EOptionParameter optionParameter)
        {
            this.OptionParameter = optionParameter;
        }

        public EOptionParameter OptionParameter { get; private set; }
    }

    public delegate void OptionsModifiedEventHandler(object sender, OptionsModifiedEventArgs e);

    [ClassInterface(ClassInterfaceType.AutoDual)]
    [CLSCompliant(false), ComVisible(true)]
    public class OptionPage : DialogPage
    {
        private EMode mode = EMode.Mru;
        private int mruCount = 5;

        public event OptionsModifiedEventHandler Modified = (s, e) => { };

        [Category("Mode")]
        [DisplayName("Startup project list population mode")]
        [Description("In Smart mode (default) projects are chosen according to their type. In MRU mode, the most recently used startup projects are displayed. ")]
        public EMode Mode
        {
            get { return mode; }
            set {
                mode = value;
                Modified(this, new OptionsModifiedEventArgs(EOptionParameter.Mode));
            }
        }

        [Category("MRU mode")]
        [DisplayName("MRU count")]
        [Description("Defines how many projects will be listed in most recently used mode. Has only effect if MRU mode is enabled.")]
        public int MruCount
        {
            get { return mruCount; }
            set { 
                mruCount = value;
                Modified(this, new OptionsModifiedEventArgs(EOptionParameter.MruCount));
            }
        }

    }
}
