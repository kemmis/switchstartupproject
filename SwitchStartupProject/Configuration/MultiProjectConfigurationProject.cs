using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace LucidConcepts.SwitchStartupProject
{
    public class MultiProjectConfigurationProject
    {
        public MultiProjectConfigurationProject(
            string nameOrPath,
            string commandLineArguments,
            string workingDirectory,
            bool startProject,
            string startExternalProgram,
            string startBrowserWithUrl)
        {
            this.NameOrPath = nameOrPath;
            this.CommandLineArguments = commandLineArguments;
            this.WorkingDirectory = workingDirectory;
            this.StartProject = startProject;
            this.StartExternalProgram = startExternalProgram;
            this.StartBrowserWithUrl = startBrowserWithUrl;
        }

        public string NameOrPath { get; private set; }
        public string CommandLineArguments { get; private set; }
        public string WorkingDirectory { get; private set; }
        public bool StartProject { get; private set; }
        public string StartExternalProgram { get; set; }
        public string StartBrowserWithUrl { get; set; }
    }
}
