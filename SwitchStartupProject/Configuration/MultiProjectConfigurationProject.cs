using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace LucidConcepts.SwitchStartupProject
{
    public class MultiProjectConfigurationProject
    {
        public MultiProjectConfigurationProject(string nameOrPath, string commandLineArguments)
        {
            this.NameOrPath = nameOrPath;
            this.CommandLineArguments = commandLineArguments;
        }

        public string NameOrPath { get; private set; }
        public string CommandLineArguments { get; private set; }
    }
}
