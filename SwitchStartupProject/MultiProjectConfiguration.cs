using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace LucidConcepts.SwitchStartupProject
{
    public class MultiProjectConfiguration
    {
        public MultiProjectConfiguration(string name, IList<string> projects)
        {
            this.Name = name;
            this.Projects = projects;
        }

        public string Name { get; private set; }
        public IList<string> Projects { get; private set; }
    }
}
