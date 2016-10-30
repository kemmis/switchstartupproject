using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace LucidConcepts.SwitchStartupProject
{
    public interface IDropdownEntry
    {
        string DisplayName { get; }
    }

    public class SingleProjectDropdownEntry : IDropdownEntry
    {
        public SingleProjectDropdownEntry(SolutionProject project)
        {
            this.Project = project;
        }

        public SolutionProject Project { get; private set; }
        public string DisplayName { get { return Project.Name; } }
    }

    public class MultiProjectDropdownEntry : IDropdownEntry
    {
        public MultiProjectDropdownEntry(MultiProjectConfiguration configuration)
        {
            this.Configuration = configuration;
        }

        public MultiProjectConfiguration Configuration { get; private set; }
        public string DisplayName { get { return Configuration.Name; } }
    }

    public class OtherDropdownEntry : IDropdownEntry
    {
        public OtherDropdownEntry(string displayName)
        {
            this.DisplayName = displayName;
        }

        public string DisplayName { get; private set; }
    }
}
