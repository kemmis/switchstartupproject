using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace LucidConcepts.SwitchStartupProject.ConfigurationsPersister
{
    public class LegacyConfigurationsProviderAdapter : IConfigurationPersister
    {
        private readonly IConfigurationPersister currentPersister;
        private readonly IList<IConfigurationPersister> legacyPersisters;

        public LegacyConfigurationsProviderAdapter(IConfigurationPersister currentPersister, params IConfigurationPersister[] legacyPersisters)
        {
            this.currentPersister = currentPersister;
            this.legacyPersisters = legacyPersisters;
        }

        public void Persist()
        {
            currentPersister.Persist();
        }

        public bool Exists(string key)
        {
            return currentPersister.Exists(key) ||
                   this.legacyPersisters.Any(legacyPersister => legacyPersister.Exists(key));
        }

        public dynamic Get(string key)
        {
            if (currentPersister.Exists(key))
                return currentPersister.Get(key);
            for (var i = legacyPersisters.Count - 1; i >= 0; i--)
            {
                if (legacyPersisters[i].Exists(key))
                    return legacyPersisters[i].Get(key);
            }
            return null;
        }

        public void Store(string key, dynamic value)
        {
            currentPersister.Store(key, value);
        }

        public bool ExistsList(string key)
        {

            return currentPersister.ExistsList(key) ||
                   this.legacyPersisters.Any(legacyPersister => legacyPersister.ExistsList(key));
        }

        public dynamic GetList(string key)
        {
            if (currentPersister.ExistsList(key))
                return currentPersister.GetList(key);
            for (var i = legacyPersisters.Count - 1; i >= 0;  i--)
            {
                if (legacyPersisters[i].ExistsList(key))
                    return legacyPersisters[i].GetList(key);
            }
            return new List<string>();
        }

        public void StoreList(string key, dynamic startupProjects)
        {
            currentPersister.StoreList(key, startupProjects);
        }

        public bool ExistsObject(string key)
        {
            return currentPersister.ExistsObject(key) ||
                   this.legacyPersisters.Any(legacyPersister => legacyPersister.ExistsObject(key));
        }

        public dynamic GetObject(string key)
        {
            if (currentPersister.ExistsObject(key))
                return currentPersister.GetObject(key);
            for (var i = legacyPersisters.Count - 1; i >= 0; i--)
            {
                if (legacyPersisters[i].ExistsObject(key))
                    return legacyPersisters[i].GetObject(key);
            }
            return null;
        }

        public void StoreObject(string key, dynamic value)
        {
            currentPersister.StoreObject(key, value);
        }

    }
}
