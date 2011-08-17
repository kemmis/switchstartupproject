// Guids.cs
// MUST match guids.h
using System;

namespace LucidConcepts.SwitchStartupProject
{
    static class GuidList
    {
        public const string guidSwitchStartupProjectPkgString = "894b6873-138b-4e5f-ac68-6863cf312f7b";
        public const string guidSwitchStartupProjectCmdSetString = "9c1b4719-1443-4d96-b264-7cf17809659b";

        public static readonly Guid guidSwitchStartupProjectCmdSet = new Guid(guidSwitchStartupProjectCmdSetString);
    };
}