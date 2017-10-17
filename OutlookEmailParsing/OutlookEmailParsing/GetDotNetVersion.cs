using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OutlookEmailParsing
{
    //This code is originally from the Microsoft Documentation
    //Please check the URL for more information
    //https://docs.microsoft.com/en-us/dotnet/framework/migration-guide/how-to-determine-which-versions-are-installed
    // Checking the version using >= will enable forward compatibility.

    public class GetDotNetVersion
    {
        public static bool Get45PlusFromRegistry()
        {
            const string subkey = @"SOFTWARE\Microsoft\NET Framework Setup\NDP\v4\Full\";

                using (RegistryKey ndpKey = RegistryKey.OpenBaseKey(RegistryHive.LocalMachine, RegistryView.Registry32).OpenSubKey(subkey))
                {
                    if (ndpKey != null && ndpKey.GetValue("Release") != null)
                    {
                        string version = CheckFor45PlusVersion((int)ndpKey.GetValue("Release"));
                        if(version.Length!= 0 && version != null)
                        {
                            return true;
                        }
                        return false;
                    }
                    else
                    {
                        Console.WriteLine(".NET Framework Version 4.5 or later is need for the utility to run. Please contact the system administrator for the installation.");
                        return false;
                    }
                }
        }
        private static string CheckFor45PlusVersion(int releaseKey)
        {
            if (releaseKey >= 460798)
                    return "4.7 or later";
                if (releaseKey >= 394802)
                    return "4.6.2";
                if (releaseKey >= 394254)
                {
                    return "4.6.1";
                }
                if (releaseKey >= 393295)
                {
                    return "4.6";
                }
                if ((releaseKey >= 379893))
                {
                    return "4.5.2";
                }
                if ((releaseKey >= 378675))
                {
                    return "4.5.1";
                }
                if ((releaseKey >= 378389))
                {
                    return "4.5";
                }
                // This code should never execute. A non-null release key should mean
                // that 4.5 or later is installed.
                return "No 4.5 or later version detected";
        }
    }
}
