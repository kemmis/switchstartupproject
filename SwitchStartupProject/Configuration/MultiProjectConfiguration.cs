using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace LucidConcepts.SwitchStartupProject
{
    public class MultiProjectConfiguration
    {
        public MultiProjectConfiguration(string name, IList<MultiProjectConfigurationProject> projects)
        {
            this.Name = name;
            this.Projects = projects;
        }

        public string Name { get; private set; }
        public IList<MultiProjectConfigurationProject> Projects { get; private set; }
    }
}
