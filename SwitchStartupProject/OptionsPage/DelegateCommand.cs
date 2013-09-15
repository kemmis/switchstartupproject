using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Input;

namespace LucidConcepts.SwitchStartupProject.OptionsPage
{
    /// <summary>
    /// A simple command that uses delegates
    /// </summary>
    public sealed class DelegateCommand : ICommand
    {
        private readonly Action executeMethod;
        private readonly Func<bool> canExecuteMethod;

        public DelegateCommand(Action executeMethod, Func<bool> canExecuteMethod = null)
        {
            this.executeMethod = executeMethod;
            this.canExecuteMethod = canExecuteMethod ?? (() => true);
        }

        public event EventHandler CanExecuteChanged = (s, e) => { };

        public void RaiseCanExecuteChanged()
        {
            CanExecuteChanged(this, EventArgs.Empty);
        }

        void ICommand.Execute(object parameter)
        {
            this.executeMethod();
        }

        bool ICommand.CanExecute(object parameter)
        {
            return this.canExecuteMethod();
        }
    }
}
