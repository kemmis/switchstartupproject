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
        All,
        Smart,
        MostRecentlyUsed,
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
        private EMode mode = EMode.All;
        private int mruCount = 5;

        public event OptionsModifiedEventHandler Modified = (s, e) => { };

        [Category("Mode")]
        [DisplayName("Choose which projects are listed")]
        [Description("All (default): All projects are listed. Smart: Projects are chosen according to their type. MostRecentlyUsed: The most recently used startup projects are listed.")]
        public EMode Mode
        {
            get { return mode; }
            set {
                mode = value;
                Modified(this, new OptionsModifiedEventArgs(EOptionParameter.Mode));
            }
        }

        [Category("Most Recently Used")]
        [DisplayName("Count")]
        [Description("Choose how many projects are listed in MostRecentlyUsed mode. Has only effect if MostRecentlyUsed mode is active.")]
        public int MostRecentlyUsedCount
        {
            get { return mruCount; }
            set { 
                mruCount = value;
                Modified(this, new OptionsModifiedEventArgs(EOptionParameter.MruCount));
            }
        }

    }
}
