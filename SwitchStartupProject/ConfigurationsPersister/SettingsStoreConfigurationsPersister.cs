using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using EnvDTE;
using Microsoft.VisualStudio.Settings;
using Microsoft.VisualStudio.Shell.Settings;

namespace LucidConcepts.SwitchStartupProject.ConfigurationsPersister
{
    public class SettingsStoreConfigurationsPersister : IConfigurationPersister
    {
        private const string collectionKeyFormat = "LucidConcepts\\SwitchStartupProject\\StartupProjectMRUListBySolution\\{0}";
        private readonly WritableSettingsStore userSettingsStore;
        private readonly DTE dte;
        private const string listKey = "MRU";

        public SettingsStoreConfigurationsPersister(IServiceProvider serviceProvider, DTE dte)
        {
            this.dte = dte;
            // get settings store
            SettingsManager settingsManager = new ShellSettingsManager(serviceProvider);
            userSettingsStore = settingsManager.GetWritableSettingsStore(SettingsScope.UserSettings);
        }

        public void Persist()
        {
        }

        public void Clean()
        {
            var collectionPath = _GetCollectionPath();
            if (userSettingsStore.CollectionExists(collectionPath))
            {
                userSettingsStore.DeleteCollection(collectionPath);
            }
        }

        public bool Exists(string key)
        {
            return false;
        }

        public string Get(string key)
        {
            throw new NotSupportedException("This legacy settings persister does not support this operation");
        }

        public void Store(string key, string value)
        {
            throw new NotSupportedException("This legacy settings persister does not support this operation");
        }


        public bool ExistsList(string key)
        {
            if (key != listKey) return false;
            var collectionPath = _GetCollectionPath();
            return userSettingsStore.CollectionExists(collectionPath);
        }

        public IEnumerable<string> GetList(string key)
        {
            if (key != listKey) throw new NotSupportedException("This legacy settings persister does not support this operation");
            var collectionPath = _GetCollectionPath();
            if (!userSettingsStore.CollectionExists(collectionPath)) return new List<string>();
            var names = userSettingsStore.GetPropertyNames(collectionPath).ToList();
            names.Sort();
            return names.Select(name => userSettingsStore.GetString(collectionPath, name));
        }



        public void StoreList(string key, IEnumerable<string> startupProjects)
        {
            if (key != listKey) throw new NotSupportedException("This legacy settings persister does not support this operation");
            var collectionPath = _GetCollectionPath();
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

        private string _GetCollectionPath()
        {
            var solutionFileName = dte.Solution.FullName;
            return string.Format(collectionKeyFormat, solutionFileName.Replace('\\', '/'));
        }


    }
}
