using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace LucidConcepts.SwitchStartupProject
{
    public class StartupConfiguration
    {
        public StartupConfiguration(string name, IList<StartupConfigurationProject> projects)
        {
            this.Name = name;
            this.Projects = projects;
        }

        public string Name { get; private set; }
        public IList<StartupConfigurationProject> Projects { get; private set; }

        public bool IsEqual(StartupConfiguration other)
        {
            return this.Name == other.Name &&
                   this.Projects.Zip(other.Projects, Tuple.Create)
                       .All(tuple => tuple.Item1.IsEqual(tuple.Item2));
        }
    }

    public class StartupConfigurationProject
    {
        public StartupConfigurationProject(
            SolutionProject project,
            string commandLineArguments,
            string workingDirectory,
            bool? startProject,
            string startExternalProgram,
            string startBrowserWithUrl,
            bool? enableRemoteDebugging,
            string remoteDebuggingMachine)
        {
            this.Project = project;
            this.CommandLineArguments = commandLineArguments;
            this.WorkingDirectory = workingDirectory;
            this.StartProject = startProject;
            this.StartExternalProgram = startExternalProgram;
            this.StartBrowserWithUrl = startBrowserWithUrl;
            this.EnableRemoteDebugging = enableRemoteDebugging;
            this.RemoteDebuggingMachine = remoteDebuggingMachine;
        }

        public SolutionProject Project { get; private set; }
        public string CommandLineArguments { get; private set; }
        public string WorkingDirectory { get; private set; }
        public bool? StartProject { get; set; }
        public string StartExternalProgram { get; set; }
        public string StartBrowserWithUrl { get; set; }
        public bool? EnableRemoteDebugging { get; private set; }
        public string RemoteDebuggingMachine { get; set; }

        public bool IsEqual(StartupConfigurationProject other)
        {
            return this.Project == other.Project &&
                   this.CommandLineArguments == other.CommandLineArguments &&
                   this.WorkingDirectory == other.WorkingDirectory &&
                   this.StartProject == other.StartProject &&
                   this.StartExternalProgram == other.StartExternalProgram &&
                   this.StartBrowserWithUrl == other.StartBrowserWithUrl &&
                   this.EnableRemoteDebugging == other.EnableRemoteDebugging &&
                   this.RemoteDebuggingMachine == other.RemoteDebuggingMachine;
        }
    }
}
