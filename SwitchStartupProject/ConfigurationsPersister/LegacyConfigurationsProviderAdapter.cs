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
            foreach (var legacyPersister in legacyPersisters)
            {
                legacyPersister.Clean();
            }
        }

        public void Clean()
        {
            currentPersister.Clean();
            legacyPersisters.Clear();
        }

        public bool Exists(string key)
        {
            return currentPersister.Exists(key) ||
                   this.legacyPersisters.Any(legacyPersister => legacyPersister.Exists(key));
        }

        public string Get(string key)
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

        public void Store(string key, string value)
        {
            currentPersister.Store(key, value);
        }

        public bool ExistsList(string key)
        {

            return currentPersister.ExistsList(key) ||
                   this.legacyPersisters.Any(legacyPersister => legacyPersister.ExistsList(key));
        }

        public IEnumerable<string> GetList(string key)
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

        public void StoreList(string key, IEnumerable<string> values)
        {
            currentPersister.StoreList(key, values);
        }
    }
}
