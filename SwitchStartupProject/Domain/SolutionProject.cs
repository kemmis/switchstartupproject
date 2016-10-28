using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

using EnvDTE;

using Microsoft.VisualStudio;
using Microsoft.VisualStudio.Shell.Interop;

namespace LucidConcepts.SwitchStartupProject
{
    public class SolutionProject
    {
        private IList<Guid> projectTypeGuids;
        private string typeName;
        private string caption;
        private Guid guid;
        private bool isWebSiteProject;
        private string solutionFullName;

        public IVsHierarchy Hierarchy { get; private set; }
        public string Name { get; private set; }
        public string Path { get; private set; }
        public Project Project { get; private set; }
        public string StartArgumentsPropertyName { get; private set; }


        protected SolutionProject(IVsHierarchy hierarchy, string name, string typeName, string caption, Guid guid, string solutionFullName)
        {
            this.Hierarchy = hierarchy;
            this.Name = name;
            this.typeName = typeName;
            this.caption = caption;
            this.guid = guid;
            this.solutionFullName = solutionFullName;

            this.Project = _GetProjectFromHierarchy(hierarchy);
            projectTypeGuids = _GetProjectTypeGuids(hierarchy, Project).ToList();
            isWebSiteProject = projectTypeGuids.Contains(GuidList.guidWebSite);

            this.Name = _GetProjectStringProperty(Hierarchy, __VSHPROPID.VSHPROPID_Name);
            this.Path = _GetProjectPath(Project, solutionFullName, isWebSiteProject);
            this.StartArgumentsPropertyName = projectTypeGuids.Contains(GuidList.guidCPlusPlus) ? "CommandArguments" : "StartArguments";
        }

        public static SolutionProject FromHierarchy(IVsHierarchy hierarchy, string solutionFullName)
        {
            var name = _GetProjectStringProperty(hierarchy, __VSHPROPID.VSHPROPID_Name);
            var typeName = _GetProjectStringProperty(hierarchy, __VSHPROPID.VSHPROPID_TypeName);
            var caption = _GetProjectStringProperty(hierarchy, __VSHPROPID.VSHPROPID_Caption);
            var guid = _GetProjectGuidProperty(hierarchy, __VSHPROPID.VSHPROPID_TypeGuid);

            // Filter out hierarchy elements that don't represent projects
            if (name == null || typeName == null || caption == null || guid == null) return null;

            return new SolutionProject(hierarchy, name, typeName, caption, guid.Value, solutionFullName);
        }

        public void Rename()
        {
            Name = _GetProjectStringProperty(Hierarchy, __VSHPROPID.VSHPROPID_Name);
            Path = _GetProjectPath(Project, solutionFullName, isWebSiteProject);
        }

        private static string _GetProjectPath(Project project, string solutionFullName, bool isWebSiteProject)
        {
            var fullPath = project.FullName;
            if (isWebSiteProject) return fullPath;  // Website projects need to be set using the full path
            var solutionPath = System.IO.Path.GetDirectoryName(solutionFullName) + @"\";
            return Paths.GetPathRelativeTo(fullPath, solutionPath);
        }

        private static string _GetProjectStringProperty(IVsHierarchy pHierarchy, __VSHPROPID property)
        {
            object valueObject = null;
            return pHierarchy.GetProperty((uint)VSConstants.VSITEMID.Root, (int)property, out valueObject) == VSConstants.S_OK ? (string)valueObject : null;
        }

        private static Guid? _GetProjectGuidProperty(IVsHierarchy pHierarchy, __VSHPROPID property)
        {
            Guid guid = Guid.Empty;
            return pHierarchy.GetGuidProperty((uint)VSConstants.VSITEMID.Root, (int)property, out guid) == VSConstants.S_OK ? guid : (Guid?)null;
        }

        private static IEnumerable<Guid> _GetProjectTypeGuids(IVsHierarchy pHierarchy, Project project)
        {
            IEnumerable<Guid> projectTypeGuids;
            var aggregatableProject = pHierarchy as IVsAggregatableProject;
            if (aggregatableProject != null)
            {
                string projectTypeGuidString;
                aggregatableProject.GetAggregateProjectTypeGuids(out projectTypeGuidString);
                projectTypeGuids = projectTypeGuidString.Split(';')
                    .Where(guidString => !string.IsNullOrEmpty(guidString))
                    .Select(guidString => new Guid(guidString));
            }
            else
            {
                projectTypeGuids = new[] { new Guid(project.Kind) };
            }
            return projectTypeGuids;
        }

        private static Project _GetProjectFromHierarchy(IVsHierarchy pHierarchy)
        {
            object project;
            ErrorHandler.ThrowOnFailure(pHierarchy.GetProperty(VSConstants.VSITEMID_ROOT,
                                                               (int)__VSHPROPID.VSHPROPID_ExtObject,
                                                               out project));
            return (project as Project);
        }

    }
}
