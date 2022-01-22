//
// Copyright (c) 2020 - 2021 by Bart Duijndam. See: https://www.duijndam.dev 
//
// Licensed under the Apache License, Version 2.0 (the "License"); 
// You may not use this file except in compliance with the License.
// You may obtain a License copy at http://www.apache.org/licenses/LICENSE-2.0
//
// Unless required by applicable law or agreed to in writing, software distributed under the License is
// distributed on an "AS IS" BASIS, WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied. 
//
// See the License for the specific language governing permissions and limitations under the License.
//
using ExcelDna.Integration;
using ExcelDna.IntelliSense;
using ExcelDna.Logging;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using System.Text;
using System.Threading.Tasks;
using Serilog;

// Added Bart
using SharpProj;
using SharpProj.Proj;

namespace TopoLib
{
    public delegate Action<ProjLogLevel, String> LOG(ProjLogLevel level, string message);

    [ComVisible(true)]
    public class AddIn : IExcelAddIn
    {
        private static ILogger _log = Serilog.Log.Logger;

        public static ILogger Logger   // property
		{
			get { return _log; }   // get method
			set { _log =  value; }  // set method
		}

        public void AutoOpen()
        {
            _log = Serilog.Log.Logger = ConfigureLogging();

            // log some info on XLL and Excel
            Version v = System.Reflection.Assembly.GetExecutingAssembly().GetName().Version;
            string version = v.ToString();
            System.DateTime date_time = System.IO.File.GetLastWriteTime(ExcelDnaUtil.XllPath);
            string compileDate = date_time.ToString();
            _log.Information($"[TOP] Starting TopoLib version {version}, compiled on {compileDate}.");

            string sBitness = Environment.Is64BitProcess ? "64-bit" : "32-bit";
            string sExcelVersion = ExcelDnaUtil.ExcelVersion.ToString();
            _log.Information($"[TOP] TopoLib running on {sBitness} Excel, version: {sExcelVersion}");

            IntelliSenseServer.Install();
        }

        public void AutoClose()
        {
            IntelliSenseServer.Uninstall();

            Serilog.Log.CloseAndFlush();
        }

        // See: https://stackoverflow.com/questions/136035/catch-multiple-exceptions-at-once
        internal static object ProcessException(Exception ex)
        {
            string errorText = "";
            
            if (CctOptions.ProjContext.LogLevel > 0)
            {
                // Deal only with exceptions thrown from within TopoLib
                // Exceptions thrown from SharpProj and/or PROJ will be caught seperately

                switch (ex)
                {
                    case NotImplementedException _:
                        errorText =  $"[TOP] {ex.GetBaseException().Message}";
                        break;

                    case ArgumentException _:
                        errorText =  $"[TOP] [{ex.GetType().Name}: {ex.Message}]";
                        break;

                    case System.Exception _:
                        errorText =  $"[TOP] {ex.GetBaseException().Message}";
                        break;


/*
                    case ArgumentNullException _:
                        errorText =  $"#:[{ex.GetBaseException().Message} argument is null]";
                        break;

                    case ProjException _:
                        errorText =  $"#:[{ex.GetBaseException().Message}]";
                        break;

                    default:
                        // you can set a breakpoint here [F9] to check if certain exception types haven't been handled yet
                        errorText =  $"#:[{ex.GetBaseException().Message}]";
                        break;
*/
                }
                AddIn.Logger.Error(errorText);
            }
    
            // This ExceptionHandler is always called from a TopoLib function or command;
            // It should return ExcelError.ExcelErrorNA to indicate an error condition

            return ExcelError.ExcelErrorNA;

        } // ProcessException

        // Create a method for the LOG delegate to handle SharProj exceptions.
        internal static void ProcessSharpProjException(ProjLogLevel level, string message)
        {
            message = "[PRO] " + message;
            int nLevel = (int)level;

            switch (nLevel)
            {
                default:
                case 0:
                    AddIn.Logger.Information(message);
                    break;
                case 1:
                    AddIn.Logger.Error(message);
                    break;
                case 2:
                    AddIn.Logger.Debug(message);
                    break;
                case 3:
                    AddIn.Logger.Verbose(message);
                    break;
            }
        } // ProcessSharpProjException

        public static void ProcessUnhandledException(Exception ex, string message = null, [CallerMemberName] string caller = null)
        {
            try
            {
                _log.Error(ex, message ?? $"[TOP] Unhandled exception on {caller}");
            }
            catch (Exception lex)
            {
                try
                {
                    Serilog.Debugging.SelfLog.WriteLine(lex.ToString());
                }
                catch
                {
                    // Do nothing...
                }
            }

            if (ex is TargetInvocationException && !(ex.InnerException is null))
            {
                ProcessUnhandledException(ex.InnerException, message, caller);
                return;
            }

#if DEBUG
            MessageBox.Show(ex.ToString(), "Unexpected Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
#else
            const string errorMessage = "An unexpected error ocurred. Please try again in a few minutes, and if the error persists, contact support";
            MessageBox.Show(errorMessage, "Unexpected Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
#endif
        }

/*
        PROJ uses the foillowing logging levels :
        public enum ProjLogLevel
        {
            None  = 0,
            Error = 1,
            Debug = 2,
            Trace = 3
        }
        In other words; for PROJ: the higher the level the MORE verbose PROJ becomes

        SERILOG uses `levels` as the primary means for assigning importance to log events. The levels in increasing order of importance are:

        Verbose     - tracing information and debugging minutiae; generally only switched on in unusual situations
        Debug       - internal control flow and diagnostic state dumps to facilitate pinpointing of recognised problems
        Information - events of interest or that have relevance to outside observers; the default enabled minimum logging level
        Warning     - indicators of possible issues or service/functionality degradation
        Error       - indicating a failure within the application or connected system
        Fatal       - critical errors causing complete failure of the application

        In other words; for SERILOG: the higher the level the LESS verbose it becomes

        The logging level can be changed dynamically by using : var levelSwitch = new LoggingLevelSwitch();
        See:  https://github.com/serilog/serilog/wiki/Writing-Log-Events#dynamic-levels

        This is not (really) required in our case; as the PROJ library will take of a changing logging level. Otherwise use the following to configure logging:

        var log = new LoggerConfiguration()
            .MinimumLevel.ControlledBy(levelSwitch)
            .WriteTo.ColoredConsole()
            .CreateLogger();
*/

        private static ILogger ConfigureLogging()
        {
            return new LoggerConfiguration()
                .MinimumLevel.Verbose()
                .WriteTo.ExcelDnaLogDisplay(displayOrder: DisplayOrder.NewestFirst)
                .CreateLogger();
        }
    }
}
