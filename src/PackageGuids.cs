using System;

namespace NF.VSTools
{
    internal static class PackageGuids
    {
        // Use the SAME GUID string as the value= in your VSCT <GuidSymbol name="guidNF_VSToolsPackageCmdSet" ...>
        public const string guidNF_VSToolsPackageCmdSetString = "{84a273b1-3686-4d23-839a-26a681df09df}";
        public static readonly Guid guidNF_VSToolsPackageCmdSet = new(guidNF_VSToolsPackageCmdSetString);
    }
}