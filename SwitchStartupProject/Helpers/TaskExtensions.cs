using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;

namespace LucidConcepts.SwitchStartupProject
{
    public static class TaskExtensions
    {
        public static void LogExceptions(this Task that, string summary)
        {
            that.ContinueWith(t => t.Exception?.Flatten().Handle(e =>
            {
                MessageBox.Show(summary + Environment.NewLine + e.Message, "SwitchStartupProject", MessageBoxButton.OK, MessageBoxImage.Exclamation);
                return true;
            }));
        }
    }
}
