using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows;

using EnvDTE;

namespace LucidConcepts.SwitchStartupProject
{
    public class ConfigurationFileOpener
    {
        private readonly DTE dte;
        private readonly string configurationFilename;
        private readonly ConfigurationLoader configurationLoader;

        public ConfigurationFileOpener(DTE dte, string configurationFilename, ConfigurationLoader configurationLoader)
        {
            this.dte = dte;
            this.configurationFilename = configurationFilename;
            this.configurationLoader = configurationLoader;
        }

        public void Open()
        {
            if (!File.Exists(configurationFilename))
            {
                try
                {
                    configurationLoader.CreateDefaultConfigurationFile();
                }
                catch (Exception e)
                {
                    MessageBox.Show("Could not create default configuration file " + configurationFilename + "\nError:\n" + e.ToString(), "SwitchStartupProject", MessageBoxButton.OK, MessageBoxImage.Exclamation);
                }
            }

            try
            {
                dte.ItemOperations.OpenFile(configurationFilename, EnvDTE.Constants.vsViewKindCode);
            }
            catch (Exception e)
            {
                MessageBox.Show("Could not open configuration file " + configurationFilename + "\nError:\n" + e.ToString(), "SwitchStartupProject", MessageBoxButton.OK, MessageBoxImage.Exclamation);
            }
        }
    }
}
