using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace LucidConcepts.SwitchStartupProject
{
    public class MultiProjectConfigurationProject
    {
        public MultiProjectConfigurationProject(string name, string commandLineArguments)
        {
            this.Name = name;
            this.CommandLineArguments = commandLineArguments;
        }

        public string Name { get; private set; }
        public string CommandLineArguments { get; private set; }
    }
}
