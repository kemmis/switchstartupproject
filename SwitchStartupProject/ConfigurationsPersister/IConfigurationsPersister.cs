using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace LucidConcepts.SwitchStartupProject.ConfigurationsPersister
{
    public interface IConfigurationPersister
    {
        /// <summary>
        /// Persist the configurations that were stored. Must be called before closing the solution.
        /// </summary>
        void Persist();

        bool Exists(string key);
        dynamic Get(string key);
        void Store(string key, dynamic value);
        
        bool ExistsList(string key);
        dynamic GetList(string key);
        void StoreList(string key, dynamic list);

        bool ExistsObject(string key);
        dynamic GetObject(string key);
        void StoreObject(string key, dynamic obj);

    }
}
