#region Copyright 2018-2021 C. Augusto Proiete & Contributors
//
// Licensed under the Apache License, Version 2.0 (the "License");
// you may not use this file except in compliance with the License.
// You may obtain a copy of the License at
//
//     http://www.apache.org/licenses/LICENSE-2.0
//
// Unless required by applicable law or agreed to in writing, software
// distributed under the License is distributed on an "AS IS" BASIS,
// WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
// See the License for the specific language governing permissions and
// limitations under the License.
//
#endregion

using ExcelDna.Integration;
using ExcelDna.Documentation;
using Serilog;

// See for more information:
// https://github.com/Excel-DNA/ExcelDna/wiki/Diagnostic-Logging
// https://github.com/serilog-contrib/serilog-sinks-exceldnalogdisplay

namespace TopoLib
{
    public static class Log
    {
        // the class name 'Log' casues a bit of a clash with Serilog's own use of 'Log'
        // the following sentence therefore doens't work (need explicit reference to SeriLog:
        // private static readonly ILogger _log = Log.Logger.ForContext(typeof(Log));

        private static readonly ILogger _log = Serilog.Log.Logger.ForContext(typeof(Log));

        [ExcelFunctionDoc(
            Name = "TL.log.Verbose",
            Description = "Writes a `Verbose` message to the LogDisplay via Serilog",
            Category = "LOG - Logging",
            HelpTopic = "TopoLib-AddIn.chm!1600",

            Returns = "The message written to the log")]
        public static string LogVerbose(string message)
        {
            // _log.Verbose(message);
            _log.Write(Serilog.Events.LogEventLevel.Verbose, message);
            return $"'[VRB] {message}' written to the log";
        }

        [ExcelFunctionDoc(
            Name = "TL.log.Debug",
            Description = "Writes a `Debug` message to the LogDisplay via Serilog",
            Category = "LOG - Logging",
            HelpTopic = "TopoLib-AddIn.chm!1601",

            Returns = "The message written to the log")]
        public static string LogDebug(string message)
        {
            // _log.Debug(message);
            _log.Write(Serilog.Events.LogEventLevel.Debug, message);
            return $"'[DBG] {message}' written to the log";
        }

/*
 * These two functions aren't supported by Proj Logging levels
 * 
        [ExcelFunctionDoc(
            Name = "TL.log.Information",
            Description = "Writes an `Information` message to the LogDisplay via Serilog",
            Category = "LOG - Logging",
            HelpTopic = "TopoLib-AddIn.chm!1602",

            Returns = "The message written to the log")]
        public static string LogInformation(string message)
        {
            _log.Information(message);
            return $"'[INF] {message}' written to the log";
        }

        [ExcelFunctionDoc(
            Name = "TL.log.Warning",
            Description = "Writes a `Warning` message to the LogDisplay via Serilog",
            Category = "LOG - Logging",
            HelpTopic = "TopoLib-AddIn.chm!1603",

            Returns = "The message written to the log")]
        public static string LogWarning(string message)
        {
            // _log.Warning(message);
            _log.Write(Serilog.Events.LogEventLevel.Warning, message);
            return $"'[WRN] {message}' written to the log";
        }
*/
        [ExcelFunctionDoc(
            Name = "TL.log.Error",
            Description = "Writes an `Error` message to the LogDisplay via Serilog",
            Category = "LOG - Logging",
            HelpTopic = "TopoLib-AddIn.chm!1604",

            Returns = "The message written to the log")]
        public static string LogError(string message)
        {
            // _log.Error(message);
            _log.Write(Serilog.Events.LogEventLevel.Error, message);
            return $"'[ERR] {message}' written to the log";
        }
    }
}
