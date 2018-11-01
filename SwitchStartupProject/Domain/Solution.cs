using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

using Microsoft.VisualStudio.Shell.Interop;

namespace LucidConcepts.SwitchStartupProject
{
    public class Solution
    {
        public Solution()
        {
            this.Projects = new Dictionary<IVsHierarchy, SolutionProject>();
            this.IsOpening = false;
        }

        public Dictionary<IVsHierarchy, SolutionProject> Projects { get; private set; }
        public ConfigurationFileTracker ConfigurationFileTracker { get; set; }
        public ConfigurationLoader ConfigurationLoader { get; set; }
        public ActiveConfigurationLoader ActiveConfigurationLoader { get; set; }
        public Configuration Configuration { get; set; }
        public bool IsOpening { get; set; }

    }
}
