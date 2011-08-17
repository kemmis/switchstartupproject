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
        MruMode,
        MruCount
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
        private bool mruMode = false;
        private int mruCount = 5;

        public event OptionsModifiedEventHandler Modified = (s, e) => { };

        [Category("MRU mode")]
        [DisplayName("Enable MRU mode")]
        [Description("If enabled, lists for each solution the 'MRU count' most recently as startup project chosen projects. If disabled, the default smart mode filters projects according to their type.")]
        public bool MruMode
        {
            get { return mruMode; }
            set { 
                mruMode = value;
                Modified(this, new OptionsModifiedEventArgs(EOptionParameter.MruMode));
            }
        }

        [Category("MRU mode")]
        [DisplayName("MRU count")]
        [Description("Defines how many most recently as startup project chosen projects will be listed. Has only effect if MRU mode is enabled.")]
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
