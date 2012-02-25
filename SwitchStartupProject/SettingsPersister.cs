using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.VisualStudio.Settings;
using Microsoft.VisualStudio.Shell.Settings;

namespace LucidConcepts.SwitchStartupProject
{
    public class SettingsPersister
    {
        private const string collectionKeyFormat = "LucidConcepts\\SwitchStartupProject\\StartupProjectMRUListBySolution\\{0}";
        private readonly WritableSettingsStore userSettingsStore;

        public SettingsPersister(IServiceProvider serviceProvider)
        {
            // get settings store
            SettingsManager settingsManager = new ShellSettingsManager(serviceProvider);
            userSettingsStore = settingsManager.GetWritableSettingsStore(SettingsScope.UserSettings);
            
        }

        public IEnumerable<string> GetMostRecentlyUsedProjectsForSolution(string solutionFileName)
        {
            var collectionPath = _GetCollectionPath(solutionFileName);
            if (!userSettingsStore.CollectionExists(collectionPath)) return new List<string>();
            var names = userSettingsStore.GetPropertyNames(collectionPath).ToList();
            names.Sort();
            return names.Select(name => userSettingsStore.GetString(collectionPath, name));
        }



        public void StoreMostRecentlyUsedProjectsForSolution(string solutionFileName, IList<string> startupProjects)
        {
            var collectionPath = _GetCollectionPath(solutionFileName);
            if (userSettingsStore.CollectionExists(collectionPath))
            {
                userSettingsStore.DeleteCollection(collectionPath);
            }
            userSettingsStore.CreateCollection(collectionPath);
            int index = 0;
            foreach (var project in startupProjects)
            {
                userSettingsStore.SetString(collectionPath, string.Format("Project{0:D2}", index++), project);
            }
        }

        private string _GetCollectionPath(string solutionFileName)
        {
            return string.Format(collectionKeyFormat, solutionFileName.Replace('\\', '/'));
        }


    }
}
