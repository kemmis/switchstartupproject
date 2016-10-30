using System;
using System.Collections.Generic;
using System.ComponentModel.Design;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;

using Microsoft.VisualStudio.Shell;

namespace LucidConcepts.SwitchStartupProject
{
    public class DropdownService
    {
        private readonly OleMenuCommand dropdownCommand;
        private const string sentinel = "";
        private const string other = "<other>";
        private const string configure = "Configure...";
        private IList<string> dropdownList;
        private string currentDropdownValue = sentinel;

        public DropdownService(IMenuCommandService mcs)
        {
            // A DropDownCombo does not let the user type into the combo box; they can only pick from the list.
            // The string value of the element selected is returned.
            // A DropDownCombo box requires two commands:

            // One command is used to ask for the current value of the combo box and to set the new value when the user makes a choice in the combo box.
            var dropodownCommandId = new CommandID(GuidList.guidSwitchStartupProjectCmdSet, (int)PkgCmdIDList.cmdidSwitchStartupProjectCombo);
            dropdownCommand = new OleMenuCommand(_HandleDropdownCommand, dropodownCommandId);
            dropdownCommand.ParametersDescription = "$"; // accept any argument string
            dropdownCommand.Enabled = false;
            mcs.AddCommand(dropdownCommand);

            // The second command is used to retrieve the list of choices for the combo box.
            var dropdownListCommandId = new CommandID(GuidList.guidSwitchStartupProjectCmdSet, (int)PkgCmdIDList.cmdidSwitchStartupProjectComboGetList);
            var dropdownListCommand = new OleMenuCommand(_HandleDropdownListCommand, dropdownListCommandId);
            mcs.AddCommand(dropdownListCommand);

            DropdownList = null;
            dropdownCommand.Enabled = true;
        }

        public static string OtherItem
        {
            get { return other; }
        }

        public string CurrentDropdownValue
        {
            get { return currentDropdownValue != sentinel ? currentDropdownValue : null; }
            set { currentDropdownValue = value ?? sentinel; }
        }

        public void ReplaceDropdownList(IList<string> listItems)
        {
            dropdownList = listItems;
            dropdownList.Add(configure);
        }

        public IList<string> DropdownList
        {
            get { return dropdownList; }
            set
            {
                dropdownList = value ?? new List<string>();
                dropdownList.Add(configure);
            }
        }

        public bool DropdownEnabled
        {
            get { return dropdownCommand.Enabled; }
            set { dropdownCommand.Enabled = value; }
        }

        public Action OnConfigurationSelected { get; set; }

        public Action<string> OnListItemSelected { get; set; }



        private void _HandleDropdownCommand(object sender, EventArgs e)
        {
            if ((null == e) || (e == EventArgs.Empty))
            {
                // We should never get here; EventArgs are required.
                throw (new ArgumentException("EventArgs are required"));
            }
            var eventArgs = e as OleMenuCmdEventArgs;
            if (eventArgs == null)
            {
                // We should never get here; EventArgs are required.
                throw (new ArgumentException("EventArgs are required"));
            }

            var newChoice = eventArgs.InValue as string;
            var currentValueHolder = eventArgs.OutValue;

            if (currentValueHolder == IntPtr.Zero && newChoice == null)
            {
                throw (new ArgumentException("Both in and out parameters should not be null"));
            }
            if (currentValueHolder != IntPtr.Zero && newChoice != null)
            {
                throw (new ArgumentException("Both in and out parameters should not be specified"));
            }

            if (currentValueHolder != IntPtr.Zero)
            {
                // when currentValueHolder is non-NULL, the IDE is requesting the current value for the combo
                Marshal.GetNativeVariantForObject(currentDropdownValue, currentValueHolder);
            }
            else if (newChoice != null)
            {
                // when newChoice is non-NULL, the IDE is sending the new value that has been selected in the combo
                _ChooseNewDropdownValue(newChoice);
            }
        }

        private void _HandleDropdownListCommand(object sender, EventArgs e)
        {
            if (e == EventArgs.Empty)
            {
                // We should never get here; EventArgs are required.
                throw (new ArgumentException("EventArgs are required"));
            }
            var eventArgs = e as OleMenuCmdEventArgs;
            if (eventArgs == null)
            {
                // We should never get here; EventArgs are required.
                throw (new ArgumentException("EventArgs are required"));
            }
            
            var inParam = eventArgs.InValue;
            var listHolder = eventArgs.OutValue;

            if (inParam != null)
            {
                throw (new ArgumentException("In parameter may not be specified"));
            }
            if (listHolder == IntPtr.Zero)
            {
                throw (new ArgumentException("Out parameter can not be NULL"));
            }
            // the second command is used to retrieve the full list of choices as an array of strings
            Marshal.GetNativeVariantForObject(dropdownList.ToArray(), listHolder);
        }

        private void _ChooseNewDropdownValue(string name)
        {
            // new value was selected
            // see if it is the configuration item
            if (String.Compare(configure, name, StringComparison.CurrentCultureIgnoreCase) == 0)
            {
                if (OnConfigurationSelected != null) OnConfigurationSelected();
                return;
            }

            // see if it is the sentinel
            if (String.Compare(sentinel, name, StringComparison.CurrentCultureIgnoreCase) == 0)
            {
                if (OnListItemSelected != null) OnListItemSelected(null);
                return;
            }

            // see if it is one of the list items
            foreach (string project in this.dropdownList)
            {
                if (String.Compare(project, name, StringComparison.CurrentCultureIgnoreCase) == 0)
                {
                    if (OnListItemSelected != null) OnListItemSelected(project);
                    return;
                }
            }
            throw (new ArgumentException("Param is not a valid string in list"));
        }

    }
}
