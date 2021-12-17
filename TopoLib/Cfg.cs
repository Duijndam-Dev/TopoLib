using System;
using System.IO;
using System.Configuration;

using System.Linq;
using System.Text;
using System.Threading.Tasks;

using ExcelDna.Integration;
using ExcelDna.Documentation;


// The purpose of this code is to define configuration settings that are serialized in a *.config file that resides in :
// Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData)
// This information can then be read from AutoOpen() to populate static variables for use with coordinate transforms
// The idea is that the default parameters can also be set from the (yet to be developed) Ribbon Interface.

namespace TopoLib
{
    public static class Cfg
    {
        // see for an example: https://docs.microsoft.com/en-us/dotnet/api/system.configuration.configurationmanager.openexeconfiguration?view=dotnet-plat-ext-6.0

        private static string getConfigFilePath()
        {
            string appData = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);
            appData = Path.Combine(appData , "TopoLib", "TopoLib") + ".config";
            return appData;
/*
 *          In case you want the Config file to sit next to the XLL-file, use this:
 *          
            string xllPath = ExcelDnaUtil.XllPath;
            string cfgPath = Path.GetDirectoryName(xllPath);
            string cfgName = Path.GetFileNameWithoutExtension(xllPath);
            cfgPath = Path.Combine(cfgPath, cfgName) + ".config";
            return cfgPath;
*/
        }

        [ExcelFunctionDoc(
            Name = "TL.cfg.AddOrUpdateKey",
            Description = "Adds new, or updates existing [key, value] pair in the TopoLib configuration file",
            Category = "CFG - Configuration",
            HelpTopic = "TopoLib-AddIn.chm!1301",

            Returns = "[key:<{key}>, value:<{value}>] string in case of succes, #VALUE in case of failure.",
            Remarks = "In case of a #VALUE error, please ensure the {key} and {value} strings are in between \"double quotes\".")]
        public static object AddOrUpdateKey
            (
            [ExcelArgument("First string of the [key, value] pair to be added/updated", Name = "key")] string key,
            [ExcelArgument("Second string of the [key, value] pair to be added/updated", Name = "value")] string value
            )
        {
            try
            {
/*
                 // Get the roaming configuration that applies to the current user.
                Configuration roamingConfig = ConfigurationManager.OpenExeConfiguration(ConfigurationUserLevel.PerUserRoamingAndLocal);
*/
                // Map the roaming configuration file. This enables the application to access
                ExeConfigurationFileMap configFileMap = new ExeConfigurationFileMap();
                configFileMap.ExeConfigFilename = getConfigFilePath();

                // Get the mapped configuration file.
                Configuration configFile = ConfigurationManager.OpenMappedExeConfiguration(configFileMap, ConfigurationUserLevel.None);
                var settings = configFile.AppSettings.Settings;

                if (settings[key] == null)
                {
                    settings.Add(key, value);
                }
                else
                {
                    settings[key].Value = value;
                }
                configFile.Save(ConfigurationSaveMode.Modified);
                ConfigurationManager.RefreshSection(configFile.AppSettings.SectionInformation.Name);

                string kv =  $"[key:<{key}>, value:<{value}>]";

                return kv;
            }
            catch (Exception ex)
            {
                return Lib.ExceptionHandler(ex);
            }
        }

        [ExcelFunctionDoc(
            Name = "TL.cfg.ClearAllKeys",
            Description = "Removes all [key, value] pairs from the TopoLib configuration file",
            Category = "CFG - Configuration",
            HelpTopic = "TopoLib-AddIn.chm!1302",

            Returns = "[all key-value pairs cleared] or [no key-value pairs present] in case of success, #VALUE in case of failure.")]
        public static object ClearAllKeys()
        {
            try
            {
                // Map the roaming configuration file. This enables the application to access
                ExeConfigurationFileMap configFileMap = new ExeConfigurationFileMap();
                configFileMap.ExeConfigFilename = getConfigFilePath();

                // Get the mapped configuration file.
                Configuration configFile = ConfigurationManager.OpenMappedExeConfiguration(configFileMap, ConfigurationUserLevel.None);
                var settings = configFile.AppSettings.Settings;

                int nRows = settings.Count;
                if (nRows > 0)
                {
                    settings.Clear();
                    configFile.Save(ConfigurationSaveMode.Modified);
                    ConfigurationManager.RefreshSection(configFile.AppSettings.SectionInformation.Name);
                    return "[all key-value pairs cleared]";
                }
                else
                    return "[no key-value pairs present]";
            }
            catch (Exception ex)
            {
                return Lib.ExceptionHandler(ex);
            }
        }

        [ExcelFunctionDoc(
            Name = "TL.cfg.GetValue",
            Description = "Gets the 'value' from a [key, value] pair from the TopoLib configuration file",
            Category = "CFG - Configuration",
            HelpTopic = "TopoLib-AddIn.chm!1303",

            Returns = "the 'value' of a [key, value] pair or [key:<{key}>, not found], #VALUE in case of failure.",
            Remarks = "In case of a #VALUE error, please ensure the {key}-string is in between \"double quotes\".")]
        public static object GetValue(
            [ExcelArgument("{key} string of the requested [key, value] pair", Name = "key")] string key )
        {
            try
            {
                // Map the roaming configuration file. This enables the application to access
                ExeConfigurationFileMap configFileMap = new ExeConfigurationFileMap();
                configFileMap.ExeConfigFilename = getConfigFilePath();

                // Get the mapped configuration file.
                Configuration configFile = ConfigurationManager.OpenMappedExeConfiguration(configFileMap, ConfigurationUserLevel.None);
                var settings = configFile.AppSettings.Settings;

                if (settings[key] != null)
                {
                    string value = settings[key].Value;
                    configFile.Save(ConfigurationSaveMode.Modified);
                    ConfigurationManager.RefreshSection(configFile.AppSettings.SectionInformation.Name);
                    return value;
                }
                else
                {
                    string kv = $"[key:<{key}>, not found]";
                    return kv;
                }
            }
            catch (Exception ex)
            {
                return Lib.ExceptionHandler(ex);
            }
        }

        [ExcelFunctionDoc(
            Name = "TL.cfg.ReadAllKeys",
            Description = "Reads all [key, value] pairs from the TopoLib configuration file",
            Category = "CFG - Configuration",
            HelpTopic = "TopoLib-AddIn.chm!1304",

            Returns = "Depending on Mode: Number of [key, value] pairs, or table with all [key and/or value] pairs from the TopoLib configuration file")]
        public static object ReadAllKeys(
            [ExcelArgument("Output mode (3); 0 = topic count, 1 = {key} list, 2 = {value} list, 3 = <{key}, {value}> list", Name = "Mode")] object mode)
        {

            int nMode = (int)Optional.Check(mode, 3.0);

            try
            {
                // Map the roaming configuration file. This enables the application to access
                ExeConfigurationFileMap configFileMap = new ExeConfigurationFileMap();
                configFileMap.ExeConfigFilename = getConfigFilePath();

                // Get the mapped configuration file.
                Configuration configFile = ConfigurationManager.OpenMappedExeConfiguration(configFileMap, ConfigurationUserLevel.None);
                var settings = configFile.AppSettings.Settings;

                int nRows = settings.Count;
                int nCols = nMode == 3 ? 2 : 1;

                if (nRows == 0)
                {
                    if (nMode == 0) return
                            (double)0.0;
                    else
                        return "[No key-value pairs found]";
                }
                else
                {
                    var res = new object[nRows, nCols];

                    int i = 0;

                    if (nMode == 0)
                    {
                        res[0, 0] = nRows;
                    }

                    if (nMode == 1)
                    {
                        foreach (var key in settings.AllKeys)
                        {
                            res[i, 0] = settings[key].Key;
                            i++;
                        }
                    }

                    if (nMode == 2)
                    {
                        foreach (var key in settings.AllKeys)
                        {
                            res[i, 0] = settings[key].Value;
                            i++;
                        }
                    }

                    if (nMode == 3)
                    {
                        foreach (var key in settings.AllKeys)
                        {
                            res[i, 0] = settings[key].Key;
                            res[i, 1] = settings[key].Value;
                            i++;
                        }
                    }
                    return res;
                }
            }
            catch (Exception ex)
            {
                return Lib.ExceptionHandler(ex);
            }
        }

        [ExcelFunctionDoc(
            Name = "TL.cfg.ReadKey",
            Description = "Reads a 'value' from [key, value] pair from the TopoLib configuration file",
            Category = "CFG - Configuration",
            HelpTopic = "TopoLib-AddIn.chm!1305",

            Returns = "[key:<{key}>, value:<{value}>] pair or [key:<{key}>, not found], #VALUE in case of failure.",
            Remarks = "In case of a #VALUE error, please ensure the {key}-string is in between \"double quotes\".")]
        public static object ReadKey(
            [ExcelArgument("{key} string of the requested [key, value] pair", Name = "key")] string key )
        {
            try
            {
                // Map the roaming configuration file. This enables the application to access
                ExeConfigurationFileMap configFileMap = new ExeConfigurationFileMap();
                configFileMap.ExeConfigFilename = getConfigFilePath();

                // Get the mapped configuration file.
                Configuration configFile = ConfigurationManager.OpenMappedExeConfiguration(configFileMap, ConfigurationUserLevel.None);
                var settings = configFile.AppSettings.Settings;

                if (settings[key] != null)
                {
                    string value = settings[key].Value;
                    configFile.Save(ConfigurationSaveMode.Modified);

                    ConfigurationManager.RefreshSection(configFile.AppSettings.SectionInformation.Name);
                    string kv = $"[key:<{key}>, value:<{value}>]";
                    return kv;
                }
                else
                {
                    string kv = $"[key:<{key}>, not found]";
                    return kv;
                }
            }
            catch (Exception ex)
            {
                return Lib.ExceptionHandler(ex);
            }
        }

        [ExcelFunctionDoc(
            Name = "TL.cfg.RemoveKey",
            Description = "Removes existing [key, value] pair from the TopoLib configuration file",
            Category = "CFG - Configuration",
            HelpTopic = "TopoLib-AddIn.chm!1306",

            Returns = "[key:<{key}> removed] or [key:<{key}>, not found] in case of succes, #VALUE in case of failure.",
            Remarks = "In case of a #VALUE error, please ensure the {key}-string is in between \"double quotes\".")]
        public static object RemoveKey
            (
            [ExcelArgument("'key' of the [key, value] pair to be removed", Name = "key")] string key
            )
        {
            try
            {
                // Map the roaming configuration file. This enables the application to access
                ExeConfigurationFileMap configFileMap = new ExeConfigurationFileMap();
                configFileMap.ExeConfigFilename = getConfigFilePath();

                // Get the mapped configuration file.
                Configuration configFile = ConfigurationManager.OpenMappedExeConfiguration(configFileMap, ConfigurationUserLevel.None);
                var settings = configFile.AppSettings.Settings;

                if (settings[key] != null)
                {
                    settings.Remove(key);
                    configFile.Save(ConfigurationSaveMode.Modified);
                    ConfigurationManager.RefreshSection(configFile.AppSettings.SectionInformation.Name);
                    string kv =  $"[key:<{key}>, removed]";
                    return kv;
                }
                else
                {
                    string kv =  $"[key:<{key}>, not found]";
                    return kv;
                }
            }
            catch (Exception ex)
            {
                return Lib.ExceptionHandler(ex);
            }
        }
    }
}
