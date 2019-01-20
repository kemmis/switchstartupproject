﻿using System;
using System.Collections.Generic;
using System.Linq;

using Microsoft.VisualStudio.Shell;

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
            string remoteDebuggingMachine,
            string profileName)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            this.Project = project;
            this.CommandLineArguments = project.EvaluateBuildMacros(commandLineArguments);
            this.WorkingDirectory = project.EvaluateBuildMacros(workingDirectory);
            this.StartProject = startProject;
            this.StartExternalProgram = project.EvaluateBuildMacros(startExternalProgram);
            this.StartBrowserWithUrl = project.EvaluateBuildMacros(startBrowserWithUrl);
            this.EnableRemoteDebugging = enableRemoteDebugging;
            this.RemoteDebuggingMachine = project.EvaluateBuildMacros(remoteDebuggingMachine);
            this.ProfileName = project.EvaluateBuildMacros(profileName);
        }

        public SolutionProject Project { get; private set; }
        public string CommandLineArguments { get; private set; }
        public string WorkingDirectory { get; private set; }
        public bool? StartProject { get; set; }
        public string StartExternalProgram { get; set; }
        public string StartBrowserWithUrl { get; set; }
        public bool? EnableRemoteDebugging { get; private set; }
        public string RemoteDebuggingMachine { get; set; }
        public string ProfileName { get; private set; }
    }
}
