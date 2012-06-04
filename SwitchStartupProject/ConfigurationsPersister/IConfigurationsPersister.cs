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

        void Clean();

        bool Exists(string key);
        string Get(string key);
        void Store(string key, string value);
        
        bool ExistsList(string key);
        IEnumerable<string> GetList(string key);
        void StoreList(string key, IEnumerable<string> list);
    }
}
