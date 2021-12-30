using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using System.Collections;
using Microsoft.Win32;

using ExcelDna.Integration;
using ExcelDna.Documentation;

// The purpose of this code is to set and get an environment variable
// This variable could (for instance) be used to set the PROJ_LIB environment variable
// PROJ_LIB defines where the database and (geotiff) grid files are located.

namespace TopoLib
{
    public static class Env
    {
        [ExcelFunctionDoc(
            IsVolatile = true,
            Name = "TL.env.SetEnvironmentVariable",
            Description = "sets a [key, value] pair in the environment settings",
            Category = "ENV - Environment",
            HelpTopic = "TopoLib-AddIn.chm!1401",

            Returns = "[env:<{key}>, value:<{value}>, mode<{mode}>] string in case of succes, #VALUE in case of failure.",
            Remarks = "In case of a #VALUE error, please ensure the {key} and {value} strings are in between \"double quotes\".")]
        public static object SetEnvironmentVariable
            (
            [ExcelArgument("First string of the [key, value] pair to be added/updated", Name = "key")] string key,
            [ExcelArgument("Second string of the [key, value] pair to be added/updated", Name = "value")] string value,
            [ExcelArgument("Option (1) ; 0 = current process, 1 = current user, 2 = current machine", Name = "mode")] object oMode)
        {
            try
            {
                int nMode   = (int)Optional.Check(oMode, 1.0);
                bool bReset = String.IsNullOrWhiteSpace(value);

                EnvironmentVariableTarget E = (EnvironmentVariableTarget) nMode;
                string mode = E.ToString();

                switch (nMode)
                {
                    case 0:
                        // The current process.
                        Environment.SetEnvironmentVariable(key, bReset ? null : value, EnvironmentVariableTarget.Process);
                        value = Environment.GetEnvironmentVariable(key, EnvironmentVariableTarget.Process) ?? "(none)";
                        break;
                    default:
                    case 1:
                        // The current user.
                        Environment.SetEnvironmentVariable(key, bReset ? null : value, EnvironmentVariableTarget.User);
                        value = Environment.GetEnvironmentVariable(key, EnvironmentVariableTarget.User) ?? "(none)";
                        break;
                    case 2:
                        // The local machine.
                        Environment.SetEnvironmentVariable(key, bReset ? null : value, EnvironmentVariableTarget.Machine);
                        value = Environment.GetEnvironmentVariable(key, EnvironmentVariableTarget.Machine) ?? "(none)";
                        break;
                }

                string kv =  $"[env:<{key}>, value:<{value}>, mode:<{mode}>]";

                return kv;
            }
            catch (Exception ex)
            {
                return Lib.ExceptionHandler(ex);
            }
        } // SetEnvironmentVariable

        [ExcelFunctionDoc(
            IsVolatile = true,
            Name = "TL.env.GetEnvironmentVariable",
            Description = "reads a [key, value] pair in the environment settings",
            Category = "ENV - Environment",
            HelpTopic = "TopoLib-AddIn.chm!1402",

            Returns = "[env:<{key}>, value:<{value}>, mode<{mode}>] string in case of succes, #VALUE in case of failure.",
            Remarks = "In case of a #VALUE error, please ensure the {key} string is in between \"double quotes\".")]
        public static object GetEnvironmentVariable
            (
            [ExcelArgument("{Key} of the [key, value] pair to be read", Name = "key")] string key,
            [ExcelArgument("Option (1) ; 0 = current process, 1 = current user, 2 = current machine", Name = "mode")] object oMode)
        {
            try
            {
                int nMode  = (int)Optional.Check(oMode, 1.0);
                string value = "";

                EnvironmentVariableTarget E = (EnvironmentVariableTarget) nMode;
                string mode = E.ToString();

                switch (nMode)
                {
                    case 0:
                        // The current process.
                        value = Environment.GetEnvironmentVariable(key, EnvironmentVariableTarget.Process) ?? "(none)";
                        break;
                    default:
                    case 1:
                        // The current user.
                        value = Environment.GetEnvironmentVariable(key, EnvironmentVariableTarget.User) ?? "(none)";
                        break;
                    case 2:
                        // The local machine.
                        value = Environment.GetEnvironmentVariable(key, EnvironmentVariableTarget.Machine) ?? "(none)";
                        break;
                }

                string kv =  $"[env:<{key}>, value:<{value}>, mode:<{mode}>]";

                return kv;
            }
            catch (Exception ex)
            {
                return Lib.ExceptionHandler(ex);
            }
        } // GetEnvironmentVariable

        [ExcelFunctionDoc(
            IsVolatile = true,
            Name = "TL.env.GetEnvironmentVariableValue",
            Description = "reads a [key, value] pair in the environment settings",
            Category = "ENV - Environment",
            HelpTopic = "TopoLib-AddIn.chm!1403",

            Returns = "Value of environment variable in case of succes, #VALUE in case of failure.",
            Remarks = "In case of a #VALUE error, please ensure the {key} string is in between \"double quotes\".")]
        public static object GetEnvironmentVariableValue
            (
            [ExcelArgument("{Key} of the [key, value] pair to be read", Name = "key")] string key,
            [ExcelArgument("Option (1) ; 0 = current process, 1 = current user, 2 = current machine", Name = "mode")] object oMode)
        {
            try
            {
                int nMode  = (int)Optional.Check(oMode, 1.0);
                string value = "";

                EnvironmentVariableTarget E = (EnvironmentVariableTarget) nMode;
                string mode = E.ToString();

                switch (nMode)
                {
                    case 0:
                        // The current process.
                        value = Environment.GetEnvironmentVariable(key, EnvironmentVariableTarget.Process) ?? "(none)";
                        break;
                    default:
                    case 1:
                        // The current user.
                        value = Environment.GetEnvironmentVariable(key, EnvironmentVariableTarget.User) ?? "(none)";
                        break;
                    case 2:
                        // The local machine.
                        value = Environment.GetEnvironmentVariable(key, EnvironmentVariableTarget.Machine) ?? "(none)";
                        break;
                }

                return value;
            }
            catch (Exception ex)
            {
                return Lib.ExceptionHandler(ex);
            }
        } // GetEnvironmentVariableValue

    }
}


