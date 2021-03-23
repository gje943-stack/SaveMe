using System;
using System.Collections.Generic;
using System.IO;
using System.Reflection;
using System.Text;

namespace src
{
    public static class VaultLocationConfig
    {
        public static string VaultLocation { get; private set; }
    
        public static void ConfigureVault()
        {
            var currentPath = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
            var parentPath = Path.GetFullPath(Path.Combine(currentPath, ".."));
            VaultLocation = $@"{parentPath}\Vault";

            if (!Directory.Exists(VaultLocation))
            {
                Directory.CreateDirectory(VaultLocation);
            }
        }
    }
}
