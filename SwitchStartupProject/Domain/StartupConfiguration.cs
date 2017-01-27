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
        public StartupConfigurationProject(SolutionProject project, string commandLineArguments, string workingDirectory, string startExternalProgram)
        {
            this.Project = project;
            this.CommandLineArguments = commandLineArguments;
            this.WorkingDirectory = workingDirectory;
            this.StartExternalProgram = startExternalProgram;
        }

        public SolutionProject Project { get; private set; }
        public string CommandLineArguments { get; private set; }
        public string WorkingDirectory { get; private set; }
        public string StartExternalProgram { get; set; }

        public bool IsEqual(StartupConfigurationProject other)
        {
            return this.Project == other.Project &&
                   this.CommandLineArguments == other.CommandLineArguments &&
                   this.WorkingDirectory == other.WorkingDirectory &&
                   this.StartExternalProgram == other.StartExternalProgram;
        }
    }
}
