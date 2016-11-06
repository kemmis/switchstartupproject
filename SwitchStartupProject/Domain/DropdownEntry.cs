using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace LucidConcepts.SwitchStartupProject
{
    public interface IDropdownEntry
    {
        string DisplayName { get; }
        bool IsEqual(IDropdownEntry other);
    }

    public class SingleProjectDropdownEntry : IDropdownEntry
    {
        public SingleProjectDropdownEntry(SolutionProject project)
        {
            this.Project = project;
        }

        public SolutionProject Project { get; private set; }
        public bool Disambiguate { get; set; }
        public string DisplayName
        {
            get
            {
                if (!Disambiguate || !Project.SolutionFolders.Any()) return Project.Name;
                return string.Format("{0} ({1})", Project.Name, string.Join("/", Project.SolutionFolders));
            }
        }

        public bool IsEqual(IDropdownEntry other)
        {
            var otherSingle = other as SingleProjectDropdownEntry;
            return otherSingle != null && this.Project == otherSingle.Project;
        }
    }

    public class MultiProjectDropdownEntry : IDropdownEntry
    {
        public MultiProjectDropdownEntry(StartupConfiguration configuration)
        {
            this.Configuration = configuration;
        }

        public StartupConfiguration Configuration { get; private set; }
        public string DisplayName { get { return Configuration.Name; } }

        public bool IsEqual(IDropdownEntry other)
        {
            var otherMulti = other as MultiProjectDropdownEntry;
            return otherMulti != null && this.Configuration.IsEqual(otherMulti.Configuration);
        }
    }

    public class OtherDropdownEntry : IDropdownEntry
    {
        public OtherDropdownEntry(string displayName)
        {
            this.DisplayName = displayName;
        }

        public string DisplayName { get; private set; }

        public bool IsEqual(IDropdownEntry other)
        {
            var otherOther = other as OtherDropdownEntry;
            return otherOther != null && this.DisplayName == otherOther.DisplayName;
        }
    }
}
