using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

using EnvDTE;

using Microsoft.VisualStudio;
using Microsoft.VisualStudio.Shell.Interop;

namespace LucidConcepts.SwitchStartupProject
{
    public class ProjectHierarchyHelper
    {
        private readonly IVsSolution2 solution;

        public ProjectHierarchyHelper(IVsSolution2 solution)
        {
            this.solution = solution;
        }

        public Project GetProjectFromHierarchy(IVsHierarchy pHierarchy)
        {
            object project;
            ErrorHandler.ThrowOnFailure(pHierarchy.GetProperty(VSConstants.VSITEMID_ROOT,
                                                               (int)__VSHPROPID.VSHPROPID_ExtObject,
                                                               out project));
            return (project as Project);
        }

        public IVsHierarchy GetHierarchyFromProject(Project project)
        {
            IVsHierarchy hierarchy;
            return this.solution.GetProjectOfUniqueName(project.UniqueName, out hierarchy) == VSConstants.S_OK ? hierarchy : null;
        }

    }
}
